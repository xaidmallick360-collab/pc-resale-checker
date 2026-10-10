#!/usr/bin/env python3
"""
PC Condition -> AI Resale Prompt Generator (Windows 10/11)

Double-click or run:  python pc_checker.py

1. A popup window asks about the things a scan can't know (country, currency,
   purchase year, looks, accessories, warranty, selling goal...).
2. The script scans the PC (specs, battery, disk health, error events, SMART)
   and reads your repair log.
3. Optional automatic checks: Wi-Fi + internet, Bluetooth, sound/camera/mic
   devices, Device Manager errors, Windows/Office license, TPM, Secure Boot,
   disk write speed.
4. Optional hands-on tests: sound (left/right/both), full-screen color and
   motion test (dead pixels, flicker), keyboard + mouse test, camera + mic,
   ports and extras checklist.
5. A result window shows ONE ready-to-paste prompt. Paste it into any AI
   (Claude, ChatGPT, Gemini...) to get the estimated resale value.

Setup
  - Python 3.8+ for Windows (includes tkinter, no extra packages)
  - Run as Administrator for full battery / event log / SMART access
  - Optional: smartmontools (smartctl) for deeper disk checks

Optional command line (skips nothing, just pre-fills the popup):
  python pc_checker.py --country Qatar --currency QAR --purchase-year 2021
  python pc_checker.py --no-gui --country Qatar --currency QAR   (no popups)

Repair log (repair_log.csv, created next to this script on first run)
  part,issue,status,date,cost,note
  status = repaired | replaced | not_working | partial
  cost   = what you paid (repaired/replaced) or ESTIMATED cost to fix
           (not_working/partial)
"""

import argparse
import csv
import datetime as dt
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile

BASE_DIR = os.path.dirname(os.path.abspath(
    sys.executable if getattr(sys, "frozen", False) else __file__))
REPAIR_LOG = os.path.join(BASE_DIR, "repair_log.csv")
REPORT_DIR = os.path.join(BASE_DIR, "reports")
DISK_PROVIDERS = {"disk", "ntfs", "stornvme", "storahci", "volmgr", "iastora"}


# --------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------
def run(cmd, timeout=180):
    try:
        r = subprocess.run(cmd, capture_output=True, text=True,
                           timeout=timeout, errors="ignore")
        return r.stdout or ""
    except Exception:
        return ""


def ps(command):
    return run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
                "-Command", command])


def ps_json(command):
    out = ps(command + " | ConvertTo-Json -Depth 3 -Compress").strip()
    if not out:
        return []
    try:
        data = json.loads(out)
    except ValueError:
        return []
    if not isinstance(data, list):
        data = [data]
    return [d for d in data if isinstance(d, dict)]


def first(rows):
    return rows[0] if rows else {}


def is_admin():
    try:
        import ctypes
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


# --------------------------------------------------------------------------
# Data collection
# --------------------------------------------------------------------------
def collect_specs():
    cs = first(ps_json("Get-CimInstance Win32_ComputerSystem | "
                       "Select Manufacturer,Model,TotalPhysicalMemory,PCSystemType"))
    prod = first(ps_json("Get-CimInstance Win32_ComputerSystemProduct | "
                         "Select Version,IdentifyingNumber"))
    cpu = first(ps_json("Get-CimInstance Win32_Processor | "
                        "Select Name,NumberOfCores,NumberOfLogicalProcessors"))
    bios = first(ps_json("Get-CimInstance Win32_BIOS | "
                         "Select SerialNumber,SMBIOSBIOSVersion"))
    os_ = first(ps_json("Get-CimInstance Win32_OperatingSystem | Select Caption"))
    gpus = [g.get("Name") for g in ps_json(
        "Get-CimInstance Win32_VideoController | Select Name") if g.get("Name")]
    disks = ps_json("Get-PhysicalDisk | Select FriendlyName,MediaType,Size,"
                    "HealthStatus,OperationalStatus")
    screen = first(ps_json("Get-CimInstance Win32_VideoController | "
                           "Select CurrentHorizontalResolution,CurrentVerticalResolution"))
    has_battery = bool(ps_json("Get-CimInstance Win32_Battery | Select Name"))

    manufacturer = (cs.get("Manufacturer") or "").strip()
    model = (cs.get("Model") or "").strip()
    version = (prod.get("Version") or "").strip()
    if "lenovo" in manufacturer.lower() and version and version.lower() != "none":
        model = version  # Lenovo keeps the friendly name here

    res = ""
    if screen.get("CurrentHorizontalResolution"):
        res = f"{screen['CurrentHorizontalResolution']}x{screen['CurrentVerticalResolution']}"

    return {
        "manufacturer": manufacturer,
        "model": model,
        "serial": (bios.get("SerialNumber") or prod.get("IdentifyingNumber") or "").strip(),
        "type": "laptop" if (cs.get("PCSystemType") == 2 or has_battery) else "desktop",
        "cpu": (cpu.get("Name") or "").strip(),
        "cores": cpu.get("NumberOfCores"),
        "threads": cpu.get("NumberOfLogicalProcessors"),
        "ram_gb": round((cs.get("TotalPhysicalMemory") or 0) / (1024 ** 3), 1),
        "gpus": gpus,
        "resolution": res,
        "disks": [{
            "name": d.get("FriendlyName"),
            "media": d.get("MediaType"),
            "size_gb": round((d.get("Size") or 0) / (1024 ** 3)),
            "health": d.get("HealthStatus"),
            "status": d.get("OperationalStatus"),
        } for d in disks],
        "os": os_.get("Caption"),
        "bios_version": bios.get("SMBIOSBIOSVersion"),
        "has_battery": has_battery,
    }


def battery_health():
    path = os.path.join(tempfile.gettempdir(), "pc_checker_battery.xml")
    try:
        os.remove(path)
    except OSError:
        pass
    run(["powercfg", "/batteryreport", "/xml", "/output", path])
    try:
        with open(path, encoding="utf-8", errors="ignore") as f:
            txt = f.read()
    except OSError:
        return {"note": "battery report unavailable (run as Administrator)"}
    d = re.search(r"<DesignCapacity>(\d+)</DesignCapacity>", txt)
    f_ = re.search(r"<FullChargeCapacity>(\d+)</FullChargeCapacity>", txt)
    if not d or not f_ or int(d.group(1)) == 0:
        return {"note": "capacity data not found"}
    design, full = int(d.group(1)), int(f_.group(1))
    return {"design_mwh": design, "full_charge_mwh": full,
            "wear_pct": round(100 * (1 - full / design), 1)}


def event_counts(days):
    rows = ps_json(
        "Get-WinEvent -FilterHashtable @{LogName='System';Level=1,2;"
        "StartTime=(Get-Date).AddDays(-%d)} -ErrorAction SilentlyContinue | "
        "Select ProviderName,Id" % days)
    out = {"days": days, "total_errors": len(rows), "hardware_whea": 0,
           "disk_errors": 0, "unexpected_shutdowns": 0, "crashes_bugcheck": 0}
    for r in rows:
        prov = r.get("ProviderName") or ""
        low = prov.lower()
        if "whea" in low:
            out["hardware_whea"] += 1
        if low in DISK_PROVIDERS:
            out["disk_errors"] += 1
        if prov == "Microsoft-Windows-Kernel-Power" and r.get("Id") == 41:
            out["unexpected_shutdowns"] += 1
        if low == "bugcheck":
            out["crashes_bugcheck"] += 1
    return out


def smart_status():
    exe = shutil.which("smartctl")
    if not exe:
        return None
    results = []
    for line in run([exe, "--scan"]).splitlines():
        args = line.split("#")[0].split()
        if not args:
            continue
        health = run([exe, "-H"] + args)
        attrs = run([exe, "-A"] + args)
        hours = pct = None
        for ln in attrs.splitlines():
            low = ln.lower()
            if "power_on_hours" in low or "power on hours" in low:
                nums = re.findall(r"[\d,]+", ln)
                if nums:
                    hours = int(nums[-1].replace(",", "") or 0)
            if "percentage used" in low:
                m = re.search(r"(\d+)\s*%", ln)
                if m:
                    pct = int(m.group(1))
        results.append({"device": args[0], "passed": "PASSED" in health.upper(),
                        "power_on_hours": hours, "ssd_percentage_used": pct})
    return results


def load_repairs():
    template = ("part,issue,status,date,cost,note\n"
                "example,Example row - ignored,example,2026-01-01,0,"
                "status = repaired | replaced | not_working | partial\n")
    if not os.path.exists(REPAIR_LOG):
        with open(REPAIR_LOG, "w", encoding="utf-8", newline="") as f:
            f.write(template)
        return []
    items = []
    with open(REPAIR_LOG, encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            status = (row.get("status") or "").strip().lower()
            if not status or status == "example":
                continue
            try:
                cost = float((row.get("cost") or "0").replace(",", ""))
            except ValueError:
                cost = 0.0
            items.append({
                "part": (row.get("part") or "").strip(),
                "issue": (row.get("issue") or "").strip(),
                "status": status,
                "date": (row.get("date") or "").strip(),
                "cost": cost,
                "note": (row.get("note") or "").strip(),
            })
    return items


# --------------------------------------------------------------------------
# Extended automatic checks: network, Bluetooth, devices, license, security
# --------------------------------------------------------------------------
LICENSE_STATUS = {0: "Unlicensed", 1: "Licensed (activated)", 2: "In grace period",
                  3: "Out-of-tolerance grace", 4: "Non-genuine grace",
                  5: "Notification mode (not activated)", 6: "Extended grace"}
OFFICE_APP_ID = "0ff1ce15-a989-479d-af46-f275c6370663"


def kv_lines(text):
    """Parse 'key : value' lines (netsh output, English Windows)."""
    out = {}
    for line in text.splitlines():
        if ":" in line:
            k, v = line.split(":", 1)
            out.setdefault(k.strip().lower(), v.strip())
    return out


def ping_host(host):
    out = run(["ping", "-n", "4", "-w", "2000", host], timeout=40)
    loss = re.search(r"\((\d+)% loss\)", out)
    avg = re.search(r"Average = (\d+)ms", out)
    return {"loss_pct": int(loss.group(1)) if loss else None,
            "avg_ms": int(avg.group(1)) if avg else None}


def check_network():
    adapters = ps_json("Get-NetAdapter -Physical | Select Name,InterfaceDescription,"
                       "Status,LinkSpeed,PhysicalMediaType")
    wifi = [a for a in adapters
            if "802.11" in str(a.get("PhysicalMediaType"))
            or "wireless" in str(a.get("InterfaceDescription")).lower()
            or "wi-fi" in str(a.get("InterfaceDescription")).lower()]
    wl = kv_lines(run(["netsh", "wlan", "show", "interfaces"]))
    drv = kv_lines(run(["netsh", "wlan", "show", "drivers"]))
    return {
        "adapters": [{"name": a.get("InterfaceDescription"), "status": a.get("Status"),
                      "speed": a.get("LinkSpeed")} for a in adapters],
        "wifi_adapter_present": bool(wifi),
        "wifi_state": wl.get("state"),
        "wifi_radio": wl.get("radio type"),
        "wifi_signal": wl.get("signal"),
        "wifi_standards": drv.get("radio types supported"),
        "internet": ping_host("8.8.8.8"),
        "dns": ping_host("google.com"),
    }


def check_bluetooth():
    devs = ps_json("Get-CimInstance Win32_PnPEntity | Where-Object { $_.PNPClass -eq "
                   "'Bluetooth' } | Select Name,Status,ConfigManagerErrorCode")
    svc = first(ps_json("Get-Service bthserv -ErrorAction SilentlyContinue | "
                        "Select @{n='State';e={[string]$_.Status}}"))
    return {"entries": len(devs),
            "errors": [d.get("Name") for d in devs
                       if d.get("ConfigManagerErrorCode") not in (0, None)],
            "service": svc.get("State")}


def check_peripherals():
    rows = ps_json("Get-CimInstance Win32_PnPEntity | Where-Object { $_.PNPClass -in "
                   "'Camera','Image','AudioEndpoint','Biometric' -or $_.Name -like "
                   "'*touch screen*' } | Select Name,PNPClass,Status,ConfigManagerErrorCode")
    out = {"camera": [], "microphone": [], "speakers": [], "biometric": [],
           "touchscreen": [], "sound": []}
    for r in rows:
        name = r.get("Name") or ""
        low = name.lower()
        cls = r.get("PNPClass") or ""
        item = {"name": name,
                "ok": r.get("ConfigManagerErrorCode") in (0, None)
                and r.get("Status") in ("OK", None)}
        if cls in ("Camera", "Image"):
            out["camera"].append(item)
        elif cls == "AudioEndpoint":
            if "microphone" in low or low.startswith("mic"):
                out["microphone"].append(item)
            elif "speaker" in low or "headphone" in low:
                out["speakers"].append(item)
        elif cls == "Biometric":
            out["biometric"].append(item)
        elif "touch screen" in low:
            out["touchscreen"].append(item)
    for d in ps_json("Get-CimInstance Win32_SoundDevice | "
                     "Select Name,Status,ConfigManagerErrorCode"):
        out["sound"].append({"name": d.get("Name") or "",
                             "ok": d.get("ConfigManagerErrorCode") in (0, None)
                             and d.get("Status") in ("OK", None)})
    return out


def check_problem_devices():
    rows = ps_json("Get-CimInstance Win32_PnPEntity | Where-Object { "
                   "$_.ConfigManagerErrorCode -ne 0 } | "
                   "Select Name,PNPClass,ConfigManagerErrorCode")
    return [{"name": r.get("Name"), "class": r.get("PNPClass"),
             "code": r.get("ConfigManagerErrorCode")} for r in rows]


def check_display():
    rows = ps_json("Get-CimInstance Win32_VideoController | Select CurrentRefreshRate")
    rates = [r.get("CurrentRefreshRate") for r in rows if r.get("CurrentRefreshRate")]
    return {"refresh_hz": rates[0] if rates else None}


def check_license():
    base = "Get-CimInstance SoftwareLicensingProduct | Where-Object { $_.PartialProductKey -and "
    sel = " } | Select Name,Description,LicenseStatus"
    win = ps_json(base + "$_.Name -like 'Windows*'" + sel)
    office = ps_json(base + "$_.ApplicationId -eq '" + OFFICE_APP_ID + "'" + sel)
    return {"windows": win, "office": office}


def check_security():
    tpm = first(ps_json("Get-CimInstance -Namespace root/cimv2/security/microsofttpm "
                        "-ClassName Win32_Tpm -ErrorAction SilentlyContinue | "
                        "Select SpecVersion"))
    sb = ps("try { Confirm-SecureBootUEFI } catch { 'unsupported' }").strip().lower()
    dom = first(ps_json("Get-CimInstance Win32_ComputerSystem | Select PartOfDomain"))
    return {
        "tpm_version": (tpm.get("SpecVersion") or "").split(",")[0].strip() or None,
        "secure_boot": {"true": "enabled", "false": "disabled"}.get(
            sb, "unknown (needs Administrator, or legacy BIOS)"),
        "domain_joined": bool(dom.get("PartOfDomain")),
    }


def disk_write_speed(mb=256):
    import time
    folder = tempfile.gettempdir()
    if shutil.disk_usage(folder).free < (mb + 512) * 1024 * 1024:
        return None
    path = os.path.join(folder, "pc_checker_speed.bin")
    chunk = os.urandom(4 * 1024 * 1024)
    t0 = time.time()
    try:
        with open(path, "wb") as f:
            for _ in range(mb // 4):
                f.write(chunk)
            f.flush()
            os.fsync(f.fileno())
        elapsed = max(time.time() - t0, 0.001)
    finally:
        try:
            os.remove(path)
        except OSError:
            pass
    return round(mb / elapsed, 1)


def collect_extra(step, tests):
    extra = {}

    def safe(key, label, fn):
        step(label)
        try:
            extra[key] = fn()
        except Exception:
            extra[key] = None

    if tests.get("auto_extra"):
        safe("network", "testing Wi-Fi, network and internet...", check_network)
        safe("bluetooth", "checking Bluetooth...", check_bluetooth)
        safe("peripherals", "checking sound, camera, mic, touch devices...",
             check_peripherals)
        safe("problem_devices", "looking for device errors...", check_problem_devices)
        safe("display", "reading display refresh rate...", check_display)
        safe("license", "verifying Windows / Office license...", check_license)
        safe("security", "checking TPM and Secure Boot...", check_security)
    if tests.get("disk_speed"):
        safe("disk_speed", "measuring disk write speed (about 10 s)...", disk_write_speed)
    return extra


def extra_lines(extra):
    """Turn the extended scan into (connectivity/device lines, license/security lines)."""
    conn, lic = [], []
    if not extra:
        return conn, lic

    n = extra.get("network")
    if n:
        if not n["adapters"]:
            conn.append("No physical network adapter detected")
        for a in n["adapters"]:
            speed = f", {a['speed']}" if a.get("speed") else ""
            conn.append(f"Network adapter: {a['name']} - {a['status']}{speed}")
        conn.append("Wi-Fi adapter: " + ("present" if n["wifi_adapter_present"] else "NOT detected"))
        if n.get("wifi_state"):
            bits = [x for x in (n.get("wifi_radio"),
                                n.get("wifi_signal") and f"signal {n['wifi_signal']}") if x]
            conn.append(f"Wi-Fi connection: {n['wifi_state']}"
                        + (f" ({', '.join(bits)})" if bits else ""))
        if n.get("wifi_standards"):
            conn.append(f"Wi-Fi standards supported by the card: {n['wifi_standards']}")
        i, d = n["internet"], n["dns"]
        if i["loss_pct"] is None:
            conn.append("Internet test (ping 8.8.8.8): FAILED or no connection during the scan")
        else:
            conn.append(f"Internet test: {i['loss_pct']}% packet loss, average {i['avg_ms']} ms")
        conn.append("DNS lookup: " + ("working" if d["loss_pct"] is not None else "FAILED"))

    b = extra.get("bluetooth")
    if b is not None:
        if b["entries"] == 0:
            conn.append("Bluetooth: NO Bluetooth device found")
        else:
            conn.append(f"Bluetooth: detected ({b['entries']} entries), "
                        f"service {b.get('service') or 'unknown'}")
        for name in b["errors"]:
            conn.append(f"Bluetooth device error: {name}")

    p = extra.get("peripherals")
    if p:
        for kind, label in (("sound", "Sound device"), ("speakers", "Speakers/headphones"),
                            ("microphone", "Microphone"), ("camera", "Camera")):
            items = p.get(kind) or []
            if not items:
                conn.append(f"{label}: not detected")
            for it in items:
                conn.append(f"{label}: {it['name']} - {'OK' if it['ok'] else 'PROBLEM'}")
        for kind, label in (("biometric", "Fingerprint/biometric"), ("touchscreen", "Touchscreen")):
            for it in p.get(kind) or []:
                conn.append(f"{label}: {it['name']} - {'OK' if it['ok'] else 'PROBLEM'}")

    pd = extra.get("problem_devices")
    if pd is not None:
        if not pd:
            conn.append("Device Manager: no devices with errors")
        else:
            conn.append(f"Device Manager: {len(pd)} device(s) with errors")
            for r in pd[:10]:
                conn.append(f"  - {r['name']} ({r['class']}), error code {r['code']}")

    disp = extra.get("display")
    if disp and disp.get("refresh_hz"):
        conn.append(f"Display refresh rate: {disp['refresh_hz']} Hz")
    if extra.get("disk_speed") is not None:
        conn.append(f"Disk sequential write speed: {extra['disk_speed']} MB/s (256 MB test)")

    lc = extra.get("license")
    if lc is not None:
        if not lc["windows"]:
            lic.append("Windows: no activated product key found (possibly NOT activated)")
        for w in lc["windows"]:
            lic.append(f"Windows: {w.get('Name')} - "
                       f"{LICENSE_STATUS.get(w.get('LicenseStatus'), w.get('LicenseStatus'))}"
                       f" - {w.get('Description')}")
        if lc["office"]:
            for o in lc["office"]:
                lic.append(f"Office: {o.get('Name')} - "
                           f"{LICENSE_STATUS.get(o.get('LicenseStatus'), o.get('LicenseStatus'))}")
        else:
            lic.append("Office: no activated Office license found")
    sec = extra.get("security")
    if sec:
        lic.append(f"TPM: {('version ' + sec['tpm_version']) if sec['tpm_version'] else 'not detected / needs Administrator'}")
        lic.append(f"Secure Boot: {sec['secure_boot']}")
        lic.append("Domain / company-managed: " + ("YES (may be locked)" if sec["domain_joined"] else "no"))
    return conn, lic


# --------------------------------------------------------------------------
# Findings and prompt
# --------------------------------------------------------------------------
def build_findings(battery, events, smart, disks):
    f = []
    if battery and battery.get("wear_pct") is not None:
        w = battery["wear_pct"]
        f.append(f"Battery wear {w}% (full-charge capacity {battery['full_charge_mwh']} "
                 f"of {battery['design_mwh']} mWh design)")
    elif battery:
        f.append(f"Battery: {battery.get('note')}")
    for d in disks:
        if (d.get("health") or "Healthy") != "Healthy":
            f.append(f"Disk '{d.get('name')}' reports health: {d.get('health')}")
    for s in smart or []:
        if not s["passed"]:
            f.append(f"SMART test FAILED on {s['device']}")
        if s.get("ssd_percentage_used") is not None:
            f.append(f"SSD {s['device']} wear: {s['ssd_percentage_used']}% used")
        if s.get("power_on_hours") is not None:
            f.append(f"Drive {s['device']} power-on hours: {s['power_on_hours']}")
    d = events["days"]
    if events["hardware_whea"]:
        f.append(f"{events['hardware_whea']} hardware (WHEA) errors in last {d} days")
    if events["disk_errors"]:
        f.append(f"{events['disk_errors']} disk/filesystem errors in last {d} days")
    if events["unexpected_shutdowns"]:
        f.append(f"{events['unexpected_shutdowns']} unexpected shutdowns in last {d} days")
    if events["crashes_bugcheck"]:
        f.append(f"{events['crashes_bugcheck']} system crashes (BSOD) in last {d} days")
    if not f:
        f.append("No faults detected by the automated scan")
    f.append(f"Total critical/error system events in last {d} days: {events['total_errors']}")
    return f


def build_prompt(specs, findings, repairs, ctx, smart_available,
                 extra=None, manual=None, skipped=None):
    na = "not provided"
    L = []
    L.append("You are an experienced used-computer appraiser. Estimate the CURRENT "
             "RESALE VALUE of the PC described below. Use the data exactly as given.")
    L.append("")
    L.append("== SELLER CONTEXT ==")
    L.append(f"Location / market: {ctx.get('country') or 'not provided (state your assumption)'}")
    L.append(f"Currency for answer: {ctx.get('currency') or 'not provided (use USD and state it)'}")
    L.append(f"Purchase year: {ctx.get('purchase_year') or 'unknown'}")
    L.append(f"Original purchase price: {ctx.get('original_price') or na}")
    L.append(f"Warranty remaining: {ctx.get('warranty') or na}")
    L.append(f"Overall cosmetic condition: {ctx.get('cosmetic') or na}")
    L.append(f"Screen condition: {ctx.get('screen') or na}")
    L.append(f"Keyboard / trackpad / ports: {ctx.get('keyboard') or na}")
    L.append(f"Included accessories: {ctx.get('accessories') or na}")
    L.append(f"Selling goal: {ctx.get('urgency') or na}")
    if ctx.get("notes"):
        L.append(f"Other notes: {ctx['notes']}")
    L.append(f"Scan date: {ctx.get('date')}")
    L.append("")
    L.append("== HARDWARE ==")
    L.append(f"Type: {specs['type']}")
    L.append(f"Brand / model: {specs['manufacturer']} {specs['model']}")
    L.append(f"Serial: {specs['serial'] or 'n/a'}")
    L.append(f"CPU: {specs['cpu']} ({specs['cores']} cores / {specs['threads']} threads)")
    L.append(f"RAM: {specs['ram_gb']} GB")
    L.append(f"GPU: {', '.join(specs['gpus']) or 'n/a'}")
    if specs["resolution"]:
        L.append(f"Display resolution: {specs['resolution']}")
    for d in specs["disks"]:
        L.append(f"Storage: {d['name']}, {d['size_gb']} GB, {d['media']}, "
                 f"health {d['health']}, status {d['status']}")
    L.append(f"OS: {specs['os']}   BIOS: {specs['bios_version']}")
    L.append("")
    L.append("== AUTOMATED HEALTH SCAN RESULTS ==")
    for item in findings:
        L.append(f"- {item}")
    if not smart_available:
        L.append("- SMART drive test was not run (smartctl not installed)")
    L.append("")
    conn, lic = extra_lines(extra)
    if conn:
        L.append("== CONNECTIVITY, DEVICES AND DISPLAY (automated) ==")
        L.extend(f"- {x}" if not x.startswith("  ") else x for x in conn)
        L.append("")
    if lic:
        L.append("== LICENSE AND SECURITY (automated) ==")
        L.extend(f"- {x}" for x in lic)
        L.append("")
    if manual:
        L.append("== HANDS-ON TEST RESULTS (reported by the seller) ==")
        L.extend(f"- {label}: {value}" for label, value in manual)
        L.append("")
    if skipped:
        L.append("== TESTS NOT PERFORMED (treat as unverified) ==")
        L.append("- " + ", ".join(skipped))
        L.append("")
    L.append("== REPAIR HISTORY ==")
    done = [r for r in repairs if r["status"] in ("repaired", "replaced")]
    if not done:
        L.append("- none recorded")
    for r in done:
        L.append(f"- {r['status'].upper()}: {r['part']} - {r['issue']} "
                 f"({r['date']}), cost {r['cost']:g}. {r['note']}".strip())
    L.append("")
    L.append("== UNRESOLVED PROBLEMS / NOT WORKING ==")
    open_items = [r for r in repairs if r["status"] in ("not_working", "partial")]
    if not open_items:
        L.append("- none recorded")
    total_open = 0.0
    for r in open_items:
        total_open += r["cost"]
        L.append(f"- {r['status'].upper()}: {r['part']} - {r['issue']}; "
                 f"estimated repair cost {r['cost']:g}. {r['note']}".strip())
    if open_items:
        L.append(f"Total estimated repair cost: {total_open:g}")
    L.append("")
    L.append("== WHAT I NEED FROM YOU ==")
    L.append("1. If you can browse the web, look up CURRENT used / sold prices for this exact "
             "model and spec in my market (local classifieds, eBay sold listings, refurbisher "
             "prices). If you cannot browse, say so and use your best knowledge, stating how "
             "recent it is.")
    L.append("2. State the market price of a healthy unit of this spec in good condition, "
             "with the sources or reasoning.")
    L.append("3. Adjust that price for every issue above (battery, disk, errors, cosmetics, "
             "screen, keyboard, age, warranty, missing accessories, repairs done, unresolved "
             "faults, Wi-Fi / Bluetooth / sound / camera problems, license status, TPM and "
             "Secure Boot (Windows 11 eligibility), locks). Show each adjustment as a line "
             "with its amount and why. Treat anything not tested as unverified and say how "
             "that affects the price.")
    L.append("4. Give final numbers in my currency: quick-sale price, fair price, optimistic "
             "price, and a minimum I should not go below. Say which one best fits my selling goal.")
    L.append("5. Add a confidence level, the top 3 things a buyer will use to negotiate, and "
             "any extra info that would make your estimate more accurate.")
    L.append("Keep the answer concise and use a table for the adjustments.")
    return "\n".join(L)


# --------------------------------------------------------------------------
# Test chooser, generic dialogs and hands-on tests (tkinter)
# --------------------------------------------------------------------------
TEST_CHOICES = [
    ("auto_extra", "Automatic checks: Wi-Fi, internet, Bluetooth, devices, license, TPM"),
    ("disk_speed", "Disk speed test (writes a 256 MB temporary file, about 10 seconds)"),
    ("sound", "Sound test (you listen: left, right, both speakers)"),
    ("screen", "Screen test (full-screen colors + motion: dead pixels, flicker)"),
    ("input", "Keyboard, mouse and touchpad test"),
    ("camera_mic", "Camera and microphone test"),
    ("ports", "Ports and extras checklist (USB, HDMI, jack, card reader, locks)"),
]
TEST_NAMES = {
    "auto_extra": "Wi-Fi / Bluetooth / device / license checks",
    "disk_speed": "disk speed test", "sound": "sound test",
    "screen": "screen test", "input": "keyboard & mouse test",
    "camera_mic": "camera & microphone test", "ports": "ports checklist",
}


def _center(win):
    win.update_idletasks()
    w, h = win.winfo_reqwidth(), win.winfo_reqheight()
    win.geometry(f"+{(win.winfo_screenwidth() - w) // 2}+{(win.winfo_screenheight() - h) // 3}")
    win.attributes("-topmost", True)
    win.after(300, lambda: win.attributes("-topmost", False))


def ask_tests():
    """Let the user tick which tests to run. Returns dict or None if cancelled."""
    import tkinter as tk
    from tkinter import ttk
    root = tk.Tk()
    root.title("PC Resale Check - Choose tests")
    root.resizable(False, False)
    frm = ttk.Frame(root, padding=16)
    frm.pack()
    ttk.Label(frm, text="Which tests do you want to run?",
              font=("Segoe UI", 12, "bold")).pack(anchor="w")
    ttk.Label(frm, text="Untick anything you want to skip. Skipped tests are marked "
                        "'unverified' in the prompt.", foreground="#666").pack(anchor="w", pady=(0, 8))
    vars_ = {}
    for key, label in TEST_CHOICES:
        vars_[key] = tk.BooleanVar(value=True)
        ttk.Checkbutton(frm, text=label, variable=vars_[key]).pack(anchor="w", pady=2)

    result = {"data": None}

    def go(_e=None):
        result["data"] = {k: v.get() for k, v in vars_.items()}
        root.destroy()

    def set_all(val):
        for v in vars_.values():
            v.set(val)

    btns = ttk.Frame(frm)
    btns.pack(fill="x", pady=(12, 0))
    ttk.Button(btns, text="Select all", command=lambda: set_all(True)).pack(side="left")
    ttk.Button(btns, text="Select none", command=lambda: set_all(False)).pack(side="left", padx=6)
    ttk.Button(btns, text="Cancel", command=root.destroy).pack(side="right", padx=(8, 0))
    ttk.Button(btns, text="Start", command=go).pack(side="right")
    root.bind("<Return>", go)
    root.bind("<Escape>", lambda e: root.destroy())
    _center(root)
    root.mainloop()
    return result["data"]


def form_dialog(title, intro, items):
    """items: [(key, label, [options])]. Returns {key: choice} or None if closed."""
    import tkinter as tk
    from tkinter import ttk
    root = tk.Tk()
    root.title(title)
    root.resizable(False, False)
    frm = ttk.Frame(root, padding=16)
    frm.pack()
    ttk.Label(frm, text=intro, justify="left").grid(columnspan=2, sticky="w", pady=(0, 10))
    vars_ = {}
    for i, (key, label, options) in enumerate(items, start=1):
        ttk.Label(frm, text=label).grid(row=i, column=0, sticky="w", pady=3, padx=(0, 12))
        v = tk.StringVar(value=options[0])
        ttk.Combobox(frm, textvariable=v, values=options, state="readonly",
                     width=36).grid(row=i, column=1, pady=3)
        vars_[key] = v
    result = {"data": None}

    def ok(_e=None):
        result["data"] = {k: v.get() for k, v in vars_.items()}
        root.destroy()

    btns = ttk.Frame(frm)
    btns.grid(row=len(items) + 1, columnspan=2, sticky="e", pady=(12, 0))
    ttk.Button(btns, text="Skip", command=root.destroy).pack(side="right", padx=(8, 0))
    ttk.Button(btns, text="OK", command=ok).pack(side="right")
    root.bind("<Return>", ok)
    root.bind("<Escape>", lambda e: root.destroy())
    _center(root)
    root.mainloop()
    return result["data"]


def check_dialog(title, intro, options):
    """Tick all that apply. Returns list of ticked options, or None if closed."""
    import tkinter as tk
    from tkinter import ttk
    root = tk.Tk()
    root.title(title)
    root.resizable(False, False)
    frm = ttk.Frame(root, padding=16)
    frm.pack()
    ttk.Label(frm, text=intro, justify="left").pack(anchor="w", pady=(0, 8))
    vars_ = []
    for opt in options:
        v = tk.BooleanVar(value=False)
        ttk.Checkbutton(frm, text=opt, variable=v).pack(anchor="w", pady=1)
        vars_.append((opt, v))
    result = {"data": None}

    def ok(_e=None):
        result["data"] = [o for o, v in vars_ if v.get()]
        root.destroy()

    ttk.Button(frm, text="Done (leave empty if no problems)", command=ok).pack(
        anchor="e", pady=(12, 0))
    root.bind("<Return>", ok)
    _center(root)
    root.mainloop()
    return result["data"]


def make_tone_wav(path, freq, left=True, right=True, secs=1.2, rate=44100):
    import math
    import struct
    import wave
    frames = bytearray()
    for i in range(int(rate * secs)):
        fade = min(1.0, i / 2000, (int(rate * secs) - i) / 2000)  # avoid clicks
        v = int(0.4 * 32767 * fade * math.sin(2 * math.pi * freq * i / rate))
        frames += struct.pack("<hh", v if left else 0, v if right else 0)
    with wave.open(path, "wb") as w:
        w.setnchannels(2)
        w.setsampwidth(2)
        w.setframerate(rate)
        w.writeframes(bytes(frames))


def sound_test():
    import tkinter as tk
    from tkinter import messagebox
    try:
        import winsound
    except ImportError:
        return []
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    results = []
    messagebox.showinfo(
        "Sound test",
        "Set the volume to about 50% and unplug any headphones.\n\n"
        "You will hear 3 tones: LEFT speaker, RIGHT speaker, then BOTH.\n"
        "Click OK to start.", parent=root)
    for name, l, r, freq in (("LEFT", True, False, 440), ("RIGHT", False, True, 660),
                             ("BOTH", True, True, 880)):
        path = os.path.join(tempfile.gettempdir(), f"pc_checker_tone_{name}.wav")
        make_tone_wav(path, freq, l, r)
        try:
            winsound.PlaySound(path, winsound.SND_FILENAME)
        finally:
            try:
                os.remove(path)
            except OSError:
                pass
        suffix = "speakers" if name == "BOTH" else "speaker only"
        heard = messagebox.askyesno(
            "Sound test", f"Did you clearly hear the tone from the {name} {suffix}?",
            parent=root)
        results.append((f"Sound - {name} speaker" if name != "BOTH" else "Sound - both speakers",
                        "heard clearly" if heard else "NOT heard / wrong side"))
    root.destroy()
    q = form_dialog("Sound test", "Last question about sound:", [
        ("quality", "Overall sound quality",
         ["Clear and loud enough", "Crackling or distorted", "Very quiet",
          "Rattling or buzzing"])])
    if q:
        results.append(("Sound - overall quality", q["quality"]))
    return results


def screen_test():
    import time
    import tkinter as tk
    root = tk.Tk()
    root.attributes("-fullscreen", True)
    root.attributes("-topmost", True)
    root.configure(cursor="none")
    canvas = tk.Canvas(root, highlightthickness=0, bg="black")
    canvas.pack(fill="both", expand=True)
    colors = [("#000000", "white"), ("#ffffff", "black"), ("#ff0000", "white"),
              ("#00ff00", "black"), ("#0000ff", "white"), ("#808080", "white")]
    st = {"i": -1, "phase": "colors", "done": False, "t0": 0.0}

    def hint(text):
        canvas.delete("hint")
        canvas.create_text(canvas.winfo_width() // 2, canvas.winfo_height() // 2, text=text,
                           fill=colors[min(max(st["i"], 0), len(colors) - 1)][1]
                           if st["phase"] == "colors" else "white",
                           font=("Segoe UI", 20), tags="hint", justify="center")
        canvas.after(3500, lambda: canvas.delete("hint"))

    def finish(_e=None):
        if st["done"]:
            return
        st["done"] = True
        root.destroy()

    def animate():
        if st["done"] or st["phase"] != "motion":
            return
        elapsed = time.time() - st["t0"]
        if elapsed > 15:
            finish()
            return
        w, h = canvas.winfo_width(), canvas.winfo_height()
        pos = (elapsed * 500) % (2 * w)
        x = pos if pos < w else 2 * w - pos
        canvas.delete("sq")
        canvas.create_rectangle(x, h // 2 - 60, x + 120, h // 2 + 60, fill="white",
                                outline="", tags="sq")
        canvas.after(16, animate)

    def nxt(_e=None):
        if st["done"]:
            return
        if st["phase"] == "colors":
            st["i"] += 1
            if st["i"] < len(colors):
                canvas.configure(bg=colors[st["i"]][0])
                canvas.delete("all")
                if st["i"] == 0:
                    hint("Click or press Space for the next color.\nEsc = finish.\n\n"
                         "Look for dead / stuck dots, lines, bright spots, color tint.")
                return
            st["phase"] = "motion"
            canvas.configure(bg="#808080")
            canvas.delete("all")
            hint("Motion check: watch the moving square for flicker, trails or tearing.\n"
                 "It ends by itself after 15 seconds (or press Esc).")
            st["t0"] = time.time()
            animate()
        else:
            finish()

    for ev in ("<Button-1>", "<space>", "<Right>", "<Return>"):
        root.bind(ev, nxt)
    root.bind("<Escape>", finish)
    root.after(150, nxt)
    root.mainloop()

    issues = check_dialog(
        "Screen test - what did you notice?",
        "Tick every problem you saw (leave empty if the screen looked perfect):",
        ["Dead or stuck pixels (dots)", "Bright spots / backlight bleed",
         "Flickering or flashing", "Horizontal or vertical lines",
         "Color tint or uneven color", "Burn-in / image retention",
         "Dim or uneven brightness", "Scratches / marks on the panel"])
    if issues is None:
        return []
    return [("Screen test", "; ".join(issues) if issues else "no problems noticed")]


KEY_ROWS = [
    ["Escape"] + [f"F{i}" for i in range(1, 13)] + ["Delete"],
    ["grave"] + list("1234567890") + ["minus", "equal", "BackSpace"],
    ["Tab"] + list("qwertyuiop") + ["bracketleft", "bracketright", "backslash"],
    ["Caps_Lock"] + list("asdfghjkl") + ["semicolon", "apostrophe", "Return"],
    ["Shift_L"] + list("zxcvbnm") + ["comma", "period", "slash", "Shift_R"],
    ["Control_L", "Win_L", "Alt_L", "space", "Alt_R", "Control_R",
     "Left", "Up", "Down", "Right"],
]
KEY_LABELS = {"Escape": "Esc", "grave": "`", "minus": "-", "equal": "=", "BackSpace": "Bksp",
              "bracketleft": "[", "bracketright": "]", "backslash": "\\", "Caps_Lock": "Caps",
              "semicolon": ";", "apostrophe": "'", "Return": "Enter", "Shift_L": "Shift",
              "Shift_R": "Shift", "comma": ",", "period": ".", "slash": "/",
              "Control_L": "Ctrl", "Control_R": "Ctrl", "Win_L": "Win", "Alt_L": "Alt",
              "Alt_R": "Alt", "space": "Space", "Delete": "Del", "Left": "<", "Right": ">",
              "Up": "^", "Down": "v"}
KEY_WIDTH = {"BackSpace": 7, "Tab": 6, "Caps_Lock": 7, "Return": 7, "Shift_L": 9,
             "Shift_R": 9, "space": 22, "Control_L": 6, "Control_R": 6}


def input_test():
    import tkinter as tk
    from tkinter import ttk
    root = tk.Tk()
    root.title("Keyboard & mouse test")
    ttk.Label(root, padding=(10, 10, 10, 4), justify="left", text=(
        "Press EVERY key on the keyboard (do not hold Shift). Keys turn green when detected.\n"
        "Then click the box below with the left, right and middle buttons and scroll.\n"
        "Click Done when finished. Keys your keyboard does not have can be ignored.")
    ).pack(anchor="w")
    labels, pressed = {}, set()
    for row in KEY_ROWS:
        fr = tk.Frame(root)
        fr.pack(anchor="w", padx=10, pady=2)
        for k in row:
            lab = tk.Label(fr, text=KEY_LABELS.get(k, k.upper() if len(k) == 1 else k),
                           width=KEY_WIDTH.get(k, 4), relief="raised", bg="#dddddd")
            lab.pack(side="left", padx=1)
            labels[k] = lab

    def on_key(e):
        k = e.keysym.lower() if len(e.keysym) == 1 else e.keysym
        if k in labels:
            labels[k].config(bg="#66cc66")
            pressed.add(k)
        return "break"

    root.bind_all("<KeyPress>", on_key)

    mouse = {"left": False, "right": False, "middle": False, "wheel": False}
    indicators = {}
    ind = tk.Frame(root)
    ind.pack(anchor="w", padx=10, pady=(10, 2))
    for name in mouse:
        lab = tk.Label(ind, text=name.upper(), width=10, relief="raised", bg="#dddddd")
        lab.pack(side="left", padx=2)
        indicators[name] = lab

    def mark(name):
        mouse[name] = True
        indicators[name].config(bg="#66cc66")

    area = tk.Label(root, text="CLICK HERE: left, right, middle button, scroll wheel\n"
                               "(touchpad: tap, two-finger tap, two-finger scroll)",
                    bg="#f4f4f4", relief="sunken", height=5)
    area.pack(fill="x", padx=10, pady=6)
    area.bind("<Button-1>", lambda e: mark("left"))
    area.bind("<Button-2>", lambda e: mark("middle"))
    area.bind("<Button-3>", lambda e: mark("right"))
    area.bind("<MouseWheel>", lambda e: mark("wheel"))

    ttk.Button(root, text="Done", command=root.destroy).pack(anchor="e", padx=10, pady=(0, 10))
    _center(root)
    root.focus_force()
    root.mainloop()

    missing = [KEY_LABELS.get(k, k) for row in KEY_ROWS for k in row if k not in pressed]
    results = []
    if not missing:
        results.append(("Keyboard - key detection", "all listed keys were detected"))
    elif len(missing) > 20:
        results.append(("Keyboard - key detection",
                        f"{len(missing)} keys not pressed/detected (test likely skipped)"))
    else:
        results.append(("Keyboard - keys not detected (may not exist on this layout)",
                        ", ".join(missing)))
    got = [n for n, v in mouse.items() if v]
    results.append(("Mouse/touchpad buttons detected", ", ".join(got) if got else "none"))
    q = form_dialog("Keyboard & touchpad", "Your own judgement:", [
        ("kb", "Keyboard overall",
         ["All keys work", "Some keys sticky or not working", "Worn or missing key legends"]),
        ("pad", "Touchpad / pointing device",
         ["Works well (move, click, scroll)", "Jumpy or ghost touches",
          "Click or scroll not working", "No touchpad (desktop)"])])
    if q:
        results.append(("Keyboard - overall", q["kb"]))
        results.append(("Touchpad", q["pad"]))
    return results


def camera_mic_test():
    import tkinter as tk
    from tkinter import messagebox
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    messagebox.showinfo("Camera test",
                        "The Camera app will open.\nCheck that you can see yourself, "
                        "then close the Camera app.", parent=root)
    try:
        os.startfile("microsoft.windows.camera:")
    except Exception:
        pass
    messagebox.showinfo("Camera test", "Click OK after you have checked the camera.",
                        parent=root)
    messagebox.showinfo("Microphone test",
                        "Sound settings will open.\nUnder 'Input', speak and watch whether "
                        "the volume bar moves.\nClose Settings when you are done.", parent=root)
    try:
        os.startfile("ms-settings:sound")
    except Exception:
        pass
    messagebox.showinfo("Microphone test", "Click OK after you have checked the microphone.",
                        parent=root)
    root.destroy()
    q = form_dialog("Camera & microphone", "What did you see / hear?", [
        ("cam", "Camera", ["Works, clear picture", "Blurry or dirty lens",
                           "Not working / not detected", "No camera"]),
        ("mic", "Microphone", ["Works, bar moves", "Very weak or noisy",
                               "Not working / not detected", "No microphone"])])
    if not q:
        return []
    return [("Camera", q["cam"]), ("Microphone", q["mic"])]


def ports_checklist():
    opts = ["Not tested", "Works", "Problem", "Not present"]
    items = [
        ("USB ports (try a USB stick in each)", opts),
        ("HDMI / display output", opts),
        ("Headphone / mic jack", opts),
        ("Charging port and charger", opts),
        ("SD card reader", opts),
        ("Fingerprint reader / face login", opts),
        ("Touchscreen", opts),
        ("Keyboard backlight", opts),
        ("Hinges, lid and body", ["Not checked", "Solid, no damage", "Loose or cracked"]),
        ("Fan noise", ["Not checked", "Quiet / normal", "Loud or rattling", "Never spins"]),
        ("BIOS / boot password set?", ["Unknown", "No", "Yes"]),
        ("Still linked to a Microsoft / Find My Device account?",
         ["Unknown", "No - signed out", "Yes"]),
        ("Company or school managed (MDM)?", ["Unknown", "No", "Yes"]),
    ]
    q = form_dialog("Ports and extras checklist",
                    "Check each item on the PC, then choose the result.\n"
                    "Leave the first option if you did not check it.",
                    [(str(i), label, o) for i, (label, o) in enumerate(items)])
    if not q:
        return []
    out = []
    for i, (label, o) in enumerate(items):
        val = q[str(i)]
        if val != o[0]:
            out.append((label.rstrip("?"), val))
    return out


def run_interactive_tests(tests):
    manual = []
    for key, fn in (("sound", sound_test), ("screen", screen_test),
                    ("input", input_test), ("camera_mic", camera_mic_test),
                    ("ports", ports_checklist)):
        if tests.get(key):
            try:
                manual.extend(fn())
            except Exception as ex:
                manual.append((TEST_NAMES[key], f"test could not run ({ex})"))
    return manual


# --------------------------------------------------------------------------
# Popup windows (tkinter)
# --------------------------------------------------------------------------
QUESTIONS = [
    # key, label, choices (None = free text)
    ("country", "Country where you will sell", None),
    ("currency", "Currency for the price (e.g. QAR, USD, EUR)", None),
    ("purchase_year", "Year you bought it", None),
    ("original_price", "Original purchase price (optional)", None),
    ("warranty", "Warranty remaining", [
        "None", "Under 6 months", "6-12 months", "Over 12 months"]),
    ("cosmetic", "Overall looks / body condition", [
        "Like new, no marks", "Light scratches or wear",
        "Visible scratches or dents", "Heavy damage or cracks"]),
    ("screen", "Screen condition (if it has one)", [
        "Perfect", "Minor marks or scratches",
        "Dead pixels, lines or backlight bleed", "Cracked"]),
    ("keyboard", "Keyboard / trackpad / ports", [
        "Everything works", "Some keys or ports worn or loose",
        "Several keys or ports not working"]),
    ("accessories", "Included accessories (charger, box, bag...)", None),
    ("urgency", "Selling goal", [
        "Quick sale (within days)", "Fair price (within weeks)",
        "Best price (I can wait)"]),
    ("notes", "Anything else a buyer should know", None),
]


def ask_questions(defaults):
    """Show the question popup. Returns a dict, or None if cancelled."""
    import tkinter as tk
    from tkinter import ttk

    root = tk.Tk()
    root.title("PC Resale Check - Questions")
    root.resizable(False, False)
    frm = ttk.Frame(root, padding=16)
    frm.grid(sticky="nsew")
    ttk.Label(frm, text="Tell me what the scan can't see",
              font=("Segoe UI", 12, "bold")).grid(columnspan=2, sticky="w", pady=(0, 10))

    vars_ = {}
    first_widget = None
    for i, (key, label, choices) in enumerate(QUESTIONS, start=1):
        ttk.Label(frm, text=label).grid(row=i, column=0, sticky="w", pady=4, padx=(0, 12))
        v = tk.StringVar(value=defaults.get(key, ""))
        vars_[key] = v
        if choices:
            w = ttk.Combobox(frm, textvariable=v, values=choices, width=42)
        else:
            w = ttk.Entry(frm, textvariable=v, width=45)
        w.grid(row=i, column=1, sticky="we", pady=4)
        first_widget = first_widget or w

    ttk.Label(frm, text="Tip: pick from the dropdown or type your own. Leave blank to skip.\n"
                    "Honest answers give a more accurate price.",
              foreground="#666", justify="left").grid(row=len(QUESTIONS) + 1, columnspan=2,
                                      sticky="w", pady=(8, 0))

    result = {"data": None}

    def start(_event=None):
        result["data"] = {k: v.get().strip() for k, v in vars_.items()}
        root.destroy()

    btns = ttk.Frame(frm)
    btns.grid(row=len(QUESTIONS) + 2, columnspan=2, sticky="e", pady=(12, 0))
    ttk.Button(btns, text="Cancel", command=root.destroy).pack(side="right", padx=(8, 0))
    ttk.Button(btns, text="Start scan", command=start).pack(side="right")
    root.bind("<Return>", start)
    root.bind("<Escape>", lambda e: root.destroy())

    root.update_idletasks()
    w, h = root.winfo_reqwidth(), root.winfo_reqheight()
    root.geometry(f"+{(root.winfo_screenwidth() - w) // 2}+{(root.winfo_screenheight() - h) // 3}")
    root.attributes("-topmost", True)
    root.after(300, lambda: root.attributes("-topmost", False))
    if first_widget:
        first_widget.focus_set()
    root.mainloop()
    return result["data"]


def confirm_not_admin():
    import tkinter as tk
    from tkinter import messagebox
    r = tk.Tk()
    r.withdraw()
    ok = messagebox.askokcancel(
        "Not running as Administrator",
        "Battery, event log and SMART data may be incomplete.\n\n"
        "Click Cancel to close, then re-run from an Administrator terminal,\n"
        "or OK to continue anyway.")
    r.destroy()
    return ok


WELCOME_TEXT = (
    "WHAT THIS TOOL DOES\n"
    "It scans this PC, asks you a few questions, and builds ONE prompt that you\n"
    "paste into any AI to get an estimated resale price.\n\n"
    "BEFORE YOU START\n"
    "  1. Plug in the charger and close other programs.\n"
    "  2. Write the repair history in repair_log.csv (button below):\n"
    "       status = repaired | replaced | not_working | partial\n"
    "       cost   = what you paid, or your estimate to fix it\n"
    "  3. Have ready: year bought, original price, what is included.\n\n"
    "WHAT HAPPENS NEXT\n"
    "  Step 1  Answer the questions window\n"
    "  Step 2  Choose which tests to run\n"
    "  Step 3  The scan runs (1-2 minutes - do not use the PC)\n"
    "  Step 4  Hands-on tests: sound, screen, keyboard, camera, ports\n"
    "          (you listen / look / press keys and report what you see)\n"
    "  Step 5  The prompt is copied - paste it into Claude, ChatGPT or Gemini\n"
)


def show_welcome():
    """Instructions popup. Returns True to continue, False to quit."""
    import tkinter as tk
    from tkinter import ttk

    root = tk.Tk()
    root.title("PC Resale Check - Instructions")
    root.resizable(False, False)
    frm = ttk.Frame(root, padding=16)
    frm.pack()
    ttk.Label(frm, text="PC Resale Check", font=("Segoe UI", 14, "bold")).pack(anchor="w")
    ttk.Label(frm, text=WELCOME_TEXT, justify="left",
              font=("Consolas", 10)).pack(anchor="w", pady=(8, 12))

    result = {"go": False}

    def go():
        result["go"] = True
        root.destroy()

    def open_log():
        load_repairs()  # creates the template file if missing
        try:
            os.startfile(REPAIR_LOG)
        except Exception:
            pass

    btns = ttk.Frame(frm)
    btns.pack(fill="x")
    ttk.Button(btns, text="Open repair_log.csv", command=open_log).pack(side="left")
    ttk.Button(btns, text="Cancel", command=root.destroy).pack(side="right", padx=(8, 0))
    ttk.Button(btns, text="Continue", command=go).pack(side="right")
    root.bind("<Return>", lambda e: go())
    root.bind("<Escape>", lambda e: root.destroy())

    root.update_idletasks()
    w, h = root.winfo_reqwidth(), root.winfo_reqheight()
    root.geometry(f"+{(root.winfo_screenwidth() - w) // 2}+{(root.winfo_screenheight() - h) // 3}")
    root.attributes("-topmost", True)
    root.after(300, lambda: root.attributes("-topmost", False))
    root.mainloop()
    return result["go"]


def run_with_progress(func):
    """Run func(update) while a 'Scanning...' window shows live step text."""
    import tkinter as tk
    from tkinter import ttk
    win = tk.Tk()
    win.title("PC Resale Check - Scanning")
    win.resizable(False, False)
    ttk.Label(win, text="Scanning this PC. Please do not use it for about a minute.",
              padding=(24, 20, 24, 6), justify="center").pack()
    status = tk.StringVar(value="Starting...")
    ttk.Label(win, textvariable=status, foreground="#0a6",
              padding=(24, 0, 24, 8)).pack()
    bar = ttk.Progressbar(win, mode="indeterminate", length=320)
    bar.pack(padx=24, pady=(0, 20))
    bar.start(12)
    win.update_idletasks()
    w, h = win.winfo_reqwidth(), win.winfo_reqheight()
    win.geometry(f"+{(win.winfo_screenwidth() - w) // 2}+{(win.winfo_screenheight() - h) // 3}")
    win.update()

    def update(msg):
        status.set(msg)
        win.update()

    try:
        return func(update)
    finally:
        win.destroy()


def copy_to_clipboard(text):
    try:
        subprocess.run(["clip"], input=text.encode("utf-16"), timeout=15)
        return True
    except Exception:
        return False


def show_result(prompt, path):
    import tkinter as tk
    from tkinter import ttk

    root = tk.Tk()
    root.title("PC Resale Check - Prompt ready")
    root.geometry("820x680")
    top = ttk.Frame(root, padding=(12, 10, 12, 4))
    top.pack(fill="x")
    ttk.Label(top, text="Your prompt is ready and copied to the clipboard.",
              font=("Segoe UI", 11, "bold")).pack(anchor="w")
    ttk.Label(top, justify="left", text=(
        "1. Open Claude, ChatGPT or Gemini (best: one that can browse the web).\n"
        "2. Click the chat box and press Ctrl+V to paste.\n"
        "3. Press Enter and read the estimated price.\n"
        "If the clipboard was overwritten, press 'Copy to clipboard' below."
    )).pack(anchor="w", pady=(4, 0))

    body = ttk.Frame(root, padding=(12, 4))
    body.pack(fill="both", expand=True)
    text = tk.Text(body, wrap="word", font=("Consolas", 10))
    sb = ttk.Scrollbar(body, command=text.yview)
    text.configure(yscrollcommand=sb.set)
    sb.pack(side="right", fill="y")
    text.pack(side="left", fill="both", expand=True)
    text.insert("1.0", prompt)

    status = tk.StringVar(value=f"Saved to: {path}")
    ttk.Label(root, textvariable=status, foreground="#555",
              padding=(12, 0)).pack(anchor="w")

    def copy_all():
        content = text.get("1.0", "end-1c")
        root.clipboard_clear()
        root.clipboard_append(content)
        root.update()
        copy_to_clipboard(content)  # also push to the Windows clipboard
        status.set("Copied to clipboard. Now paste it into your AI.")

    def open_folder():
        try:
            os.startfile(os.path.dirname(path))
        except Exception:
            pass

    bar = ttk.Frame(root, padding=12)
    bar.pack(fill="x")
    ttk.Button(bar, text="Close", command=root.destroy).pack(side="right")
    ttk.Button(bar, text="Open reports folder", command=open_folder).pack(side="right", padx=8)
    ttk.Button(bar, text="Copy to clipboard", command=copy_all).pack(side="right")

    copy_all()  # copy automatically on open
    root.mainloop()


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------
def do_scan(days, progress=lambda m: None, tests=None):
    tests = tests or {"auto_extra": True, "disk_speed": False}
    total = 5 + (7 if tests.get("auto_extra") else 0) + (1 if tests.get("disk_speed") else 0)
    n = {"i": 0}

    def step(msg):
        n["i"] += 1
        progress(f"Step {n['i']}/{total}: {msg}")

    step("reading hardware details...")
    specs = collect_specs()
    step("checking battery health...")
    battery = battery_health() if specs["has_battery"] else None
    step(f"reading {days} days of error events...")
    events = event_counts(days)
    step("testing disk health (SMART)...")
    smart = smart_status()
    step("reading your repair log...")
    repairs = load_repairs()
    extra = collect_extra(step, tests) if (tests.get("auto_extra") or tests.get("disk_speed")) else None
    return specs, battery, events, smart, repairs, extra


def main():
    ap = argparse.ArgumentParser(description="Build an AI-ready PC resale appraisal prompt")
    ap.add_argument("--country", default="")
    ap.add_argument("--currency", default="")
    ap.add_argument("--purchase-year", default="")
    ap.add_argument("--cosmetic", default="")
    ap.add_argument("--accessories", default="")
    ap.add_argument("--days", type=int, default=30, help="event log window in days")
    ap.add_argument("--no-gui", action="store_true", help="no popups, console only")
    args = ap.parse_args()

    if os.name != "nt":
        sys.exit("This script is for Windows.")

    ctx = {
        "country": args.country, "currency": args.currency,
        "purchase_year": args.purchase_year, "cosmetic": args.cosmetic,
        "accessories": args.accessories,
    }

    use_gui = not args.no_gui
    if use_gui:
        try:
            import tkinter  # noqa: F401
        except ImportError:
            print("[!] tkinter not available - continuing without popups.")
            use_gui = False

    if use_gui:
        if not is_admin() and not confirm_not_admin():
            return
        if not show_welcome():
            return
        answers = ask_questions(ctx)
        if answers is None:
            return  # cancelled
        ctx.update(answers)
        tests = ask_tests()
        if tests is None:
            return
        specs, battery, events, smart, repairs, extra = run_with_progress(
            lambda upd: do_scan(args.days, upd, tests))
        manual = run_interactive_tests(tests)
    else:
        if not is_admin():
            print("[!] Not Administrator: battery/event/SMART data may be incomplete.\n")
        print("Tip: fill in repair_log.csv first for a more accurate price.\n")
        tests = {"auto_extra": True, "disk_speed": False}
        specs, battery, events, smart, repairs, extra = do_scan(args.days, print, tests)
        manual = []
        print("\n(Hands-on tests - sound, screen, keyboard, camera, ports - need the "
              "popup windows. Run without --no-gui to use them.)")

    ctx["date"] = dt.datetime.now().strftime("%Y-%m-%d")
    findings = build_findings(battery, events, smart, specs["disks"])
    skipped = [TEST_NAMES[k] for k, _ in TEST_CHOICES if not tests.get(k)]
    prompt = build_prompt(specs, findings, repairs, ctx, smart is not None,
                          extra, manual, skipped)

    os.makedirs(REPORT_DIR, exist_ok=True)
    serial = re.sub(r"[^A-Za-z0-9_-]", "", specs["serial"]) or "pc"
    path = os.path.join(REPORT_DIR, f"{serial}_{dt.datetime.now():%Y%m%d_%H%M%S}_prompt.txt")
    with open(path, "w", encoding="utf-8") as f:
        f.write(prompt)

    if use_gui:
        show_result(prompt, path)
    else:
        print("\n" + "=" * 60 + "\n" + prompt + "\n" + "=" * 60)
        print("\nSaved:", path)
        print("Copied to clipboard." if copy_to_clipboard(prompt)
              else "Open the saved file and copy its content.")
        print("\nNEXT: open Claude / ChatGPT / Gemini, press Ctrl+V in the chat box, "
              "press Enter.")


if __name__ == "__main__":
    main()
