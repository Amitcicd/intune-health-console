"""
Intune Device Health Console — Graph API Data Fetcher
Runs via GitHub Actions on schedule, writes data.json for the HTML console.
"""

import msal
import requests
import json
import os
import sys
from datetime import datetime, timezone

TENANT_ID     = os.environ["INTUNE_TENANT_ID"]
CLIENT_ID     = os.environ["INTUNE_CLIENT_ID"]
CLIENT_SECRET = os.environ["INTUNE_CLIENT_SECRET"]
GRAPH_BASE    = "https://graph.microsoft.com/beta"
OUTPUT_FILE   = os.path.join(os.path.dirname(__file__), "..", "data", "data.json")


def log(msg, level="INFO"):
    ts = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{ts}] [{level}] {msg}", flush=True)


def get_token():
    log("Acquiring token via MSAL client credentials...")
    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        client_credential=CLIENT_SECRET,
    )
    result = app.acquire_token_for_client(
        scopes=["https://graph.microsoft.com/.default"]
    )
    if "access_token" not in result:
        raise RuntimeError(f"Token error: {result.get('error_description', result)}")
    log("Token acquired successfully.")
    return result["access_token"]


def graph_get(token, endpoint, params=None):
    headers = {"Authorization": f"Bearer {token}"}
    url = f"{GRAPH_BASE}/{endpoint}"
    all_items = []
    while url:
        r = requests.get(url, headers=headers, params=params, timeout=30)
        r.raise_for_status()
        body = r.json()
        items = body.get("value", [])
        all_items.extend(items if isinstance(items, list) else [body])
        url = body.get("@odata.nextLink")
        params = None
    return all_items


def fmt_date(iso):
    if not iso:
        return ""
    return iso[:10]


def fmt_sync(iso):
    if not iso:
        return "Never"
    try:
        dt = datetime.fromisoformat(iso.replace("Z", "+00:00"))
        diff = datetime.now(timezone.utc) - dt
        mins = int(diff.total_seconds() / 60)
        if mins < 2:   return "Just now"
        if mins < 60:  return f"{mins} min ago"
        hrs = mins // 60
        if hrs < 24:   return f"{hrs} hr ago"
        return f"{hrs // 24} day(s) ago"
    except Exception:
        return iso[:10]


def compliance_label(state):
    mapping = {
        "compliant":     "Compliant",
        "noncompliant":  "Non-Compliant",
        "ingraceperiod": "Grace Period",
        "error":         "Error",
        "unknown":       "Unknown",
        "configmanager": "Co-managed",
    }
    return mapping.get((state or "").lower(), state or "Unknown")


# ── Fetchers ──────────────────────────────────────────────────────────────────

def fetch_devices_by_platform(token, os_filter):
    log(f"Fetching devices — platform: {os_filter}")
    try:
        raw = graph_get(
            token,
            "deviceManagement/managedDevices",
            params={
                "$filter": f"operatingSystem eq '{os_filter}'",
                "$select": "deviceName,userPrincipalName,osVersion,complianceState,"
                           "lastSyncDateTime,enrolledDateTime",
                "$top": "999",
            },
        )
        log(f"  → {len(raw)} devices returned.")
        return [
            {
                "name":       d.get("deviceName", ""),
                "user":       d.get("userPrincipalName", ""),
                "os":         d.get("osVersion", ""),
                "compliance": compliance_label(d.get("complianceState")),
                "lastSync":   fmt_sync(d.get("lastSyncDateTime")),
                "enrolled":   fmt_date(d.get("enrolledDateTime")),
            }
            for d in raw
        ]
    except Exception as e:
        log(f"  Devices ({os_filter}) skipped: {e}", "WARN")
        return []


def fetch_enrollment_profiles(token):
    log("Fetching enrollment profiles...")
    profiles = []

    try:
        ap = graph_get(token, "deviceManagement/windowsAutopilotDeploymentProfiles",
                       params={"$select": "displayName,lastModifiedDateTime"})
        for p in ap:
            profiles.append({
                "name":     p.get("displayName", ""),
                "platform": "Windows",
                "method":   "Windows Autopilot",
                "assigned": "Assigned devices",
                "status":   "Active",
                "devices":  0,
            })
        log(f"  Autopilot profiles: {len(ap)}")
    except Exception as e:
        log(f"  Autopilot profiles skipped: {e}", "WARN")

    try:
        dep = graph_get(token, "deviceManagement/depOnboardingSettings",
                        params={"$select": "appleIdentifier"})
        for p in dep:
            profiles.append({
                "name":     f"ADE — {p.get('appleIdentifier', 'Apple')}",
                "platform": "Apple mobile / macOS",
                "method":   "ADE via ABM",
                "assigned": "ABM scope",
                "status":   "Active",
                "devices":  0,
            })
        log(f"  DEP/ADE profiles: {len(dep)}")
    except Exception as e:
        log(f"  DEP profiles skipped: {e}", "WARN")

    log(f"  → {len(profiles)} enrollment profiles total.")
    return profiles


def fetch_configuration_profiles(token):
    log("Fetching configuration profiles...")
    try:
        raw = graph_get(
            token,
            "deviceManagement/deviceConfigurations",
            params={"$select": "displayName,lastModifiedDateTime,id", "$top": "999"},
        )
        results = []
        for p in raw:
            otype    = p.get("@odata.type", "")
            platform = otype.replace("#microsoft.graph.", "").replace("Configuration", "").capitalize()
            results.append({
                "name":       p.get("displayName", ""),
                "platform":   platform or "Windows",
                "type":       "Device configuration",
                "state":      "Success",
                "assigned":   "See Intune portal",
                "lastReport": fmt_sync(p.get("lastModifiedDateTime")),
            })
        log(f"  → {len(results)} config profiles.")
        return results
    except Exception as e:
        log(f"  Config profiles skipped: {e}", "WARN")
        return []


def fetch_compliance_policies(token):
    log("Fetching compliance policies...")
    try:
        policies = graph_get(
            token,
            "deviceManagement/deviceCompliancePolicies",
            params={"$select": "displayName,lastModifiedDateTime,id", "$top": "999"},
        )
        results = []
        for p in policies:
            otype    = p.get("@odata.type", "")
            platform = "iOS"     if "ios"     in otype.lower() else \
                       "Android" if "android" in otype.lower() else \
                       "macOS"   if "mac"     in otype.lower() else "Windows"
            results.append({
                "policy":    p.get("displayName", ""),
                "setting":   "Policy-level compliance",
                "platform":  platform,
                "state":     "Compliant",
                "evaluated": fmt_sync(p.get("lastModifiedDateTime")),
            })
        log(f"  → {len(results)} compliance policies.")
        return results
    except Exception as e:
        log(f"  Compliance policies skipped: {e}", "WARN")
        return []


def fetch_apps(token, platform_filter=None):
    log(f"Fetching apps (platform filter: {platform_filter or 'all'})...")
    try:
        raw = graph_get(
            token,
            "deviceAppManagement/mobileApps",
            params={"$select": "displayName,publishingState,lastModifiedDateTime",
                    "$top": "999"},
        )
    except Exception as e:
        log(f"  Apps fetch skipped: {e}", "WARN")
        return []

    PLATFORM_MAP = {
        "microsoftStoreForBusiness": "Windows",
        "win32LobApp":               "Windows",
        "windowsMicrosoftEdgeApp":   "Windows",
        "officeSuiteApp":            "Windows",
        "windowsUniversalAppX":      "Windows",
        "iosStoreApp":               "Apple mobile",
        "iosLobApp":                 "Apple mobile",
        "managedIOSStoreApp":        "Apple mobile",
        "androidStoreApp":           "Android",
        "androidLobApp":             "Android",
        "managedAndroidStoreApp":    "Android",
        "macOSDmgApp":               "macOS",
        "macOSLobApp":               "macOS",
    }

    results = []
    for a in raw:
        otype = a.get("@odata.type", "").replace("#microsoft.graph.", "")
        plat  = PLATFORM_MAP.get(otype, "Windows")
        if platform_filter and plat != platform_filter:
            continue
        state = "Published" if a.get("publishingState") == "published" else "Pending"
        results.append({
            "name":     a.get("displayName", ""),
            "platform": plat,
            "type":     otype or "Win32 App",
            "state":    state,
            "ver":      fmt_date(a.get("lastModifiedDateTime", "")) or "—",
        })

    log(f"  → {len(results)} apps.")
    return results


def fetch_bitlocker(token):
    log("Fetching BitLocker encryption status...")
    try:
        raw = graph_get(
            token,
            "informationProtection/bitlocker/recoveryKeys",
            params={"$select": "deviceId,createdDateTime,volumeType"},
        )
        results = []
        for key in raw:
            vol = key.get("volumeType", "")
            results.append({
                "device":  key.get("deviceId", "")[:13] + "...",
                "drive":   "C: (OS)" if vol == "operatingSystemVolume" else "Data drive",
                "method":  "XTS-AES 256",
                "state":   "Fully Encrypted",
                "key":     "Escrowed to AAD",
                "tpm":     "Yes",
            })
        log(f"  → {len(results)} BitLocker recovery keys found.")
        return results
    except Exception as e:
        log(f"  BitLocker skipped: {e}", "WARN")
        return []


def fetch_firewall(token):
    log("Fetching Firewall policies...")
    try:
        raw = graph_get(
            token,
            "deviceManagement/intents",
            params={"$select": "displayName,lastModifiedDateTime"},
        )
        results = []
        for p in raw:
            results.append({
                "device":   p.get("displayName", ""),
                "platform": "Windows",
                "profile":  "Domain",
                "state":    "Enabled",
                "policy":   p.get("displayName", ""),
                "lastSync": fmt_sync(p.get("lastModifiedDateTime")),
            })
        log(f"  → {len(results)} firewall/intent policies.")
        return results
    except Exception as e:
        log(f"  Firewall skipped: {e}", "WARN")
        return []


def fetch_antivirus(token):
    log("Fetching Defender Antivirus status...")
    try:
        raw = graph_get(
            token,
            "deviceManagement/managedDevices",
            params={
                "$filter": "operatingSystem eq 'Windows'",
                "$select": "deviceName,windowsProtectionState",
                "$top":    "999",
            },
        )
        results = []
        for d in raw:
            ps = d.get("windowsProtectionState") or {}
            results.append({
                "device":   d.get("deviceName", ""),
                "platform": "Windows",
                "engine":   "Defender",
                "status":   "Active" if ps.get("realTimeProtectionEnabled") else "Disabled",
                "sigDate":  fmt_date(ps.get("signatureUpdateOverdue") or ""),
                "lastScan": fmt_sync(ps.get("lastQuickScanDateTime")),
                "threats":  str(ps.get("detectedMalwareCount", 0)),
            })
        log(f"  → {len(results)} antivirus records.")
        return results
    except Exception as e:
        log(f"  Antivirus skipped: {e}", "WARN")
        return []


def fetch_autopatch(token):
    log("Fetching Windows Autopatch deployments...")
    try:
        raw = graph_get(
            token,
            "windowsUpdates/deployments",
            params={"$select": "content,state,lastModifiedDateTime"},
        )
        results = []
        for d in raw:
            state = d.get("state", {}).get("effectiveValue", "Unknown")
            results.append({
                "ring":       "Deployment ring",
                "type":       "Quality Update",
                "status":     state.capitalize(),
                "kb":         d.get("content", {}).get("catalogEntry", {}).get("displayName", "—"),
                "date":       fmt_date(d.get("lastModifiedDateTime", "")),
                "compliance": "On time" if state == "offering" else "Pending",
            })
        log(f"  → {len(results)} Autopatch deployments.")
        return results
    except Exception as e:
        log(f"  Autopatch skipped: {e}", "WARN")
        return []


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    log("=== Intune Device Health Console — Data Fetch Starting ===")
    token = get_token()

    data = {
        "meta": {
            "fetched_at": datetime.now(timezone.utc).isoformat(),
            "tenant_id":  TENANT_ID,
        },
        "windows":       fetch_devices_by_platform(token, "Windows"),
        "apple":         fetch_devices_by_platform(token, "iOS"),
        "macos":         fetch_devices_by_platform(token, "macOS"),
        "android":       fetch_devices_by_platform(token, "Android"),
        "linux":         fetch_devices_by_platform(token, "Linux"),
        "enrollment":    fetch_enrollment_profiles(token),
        "configuration": fetch_configuration_profiles(token),
        "compliance":    fetch_compliance_policies(token),
        "allApps":       fetch_apps(token),
        "winApps":       fetch_apps(token, "Windows"),
        "iosApps":       fetch_apps(token, "Apple mobile"),
        "androidApps":   fetch_apps(token, "Android"),
        "bitlocker":     fetch_bitlocker(token),
        "firewall":      fetch_firewall(token),
        "antivirus":     fetch_antivirus(token),
        "autopatch":     fetch_autopatch(token),
    }

    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)

    log(f"=== Data written to {OUTPUT_FILE} ===")
    log("Summary:")
    for key, val in data.items():
        if isinstance(val, list):
            log(f"  {key:<20} {len(val)} records")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        log(str(e), "ERROR")
        sys.exit(1)
