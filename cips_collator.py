#!/usr/bin/env python3
"""
CIPS Auto Collator — Full Automation (Local Edition)
=====================================================
Fetches pending tickets, collates CSVs, merges into Master,
deduplicates, raises/updates internal tickets, resolves source tickets,
IMP check via browser bridge, generates pending sheets.

Usage: Double-click "Run CIPS Collator.bat"
"""

import csv, io, json, os, re, shutil, sys, time, webbrowser
from collections import defaultdict
from datetime import datetime, timedelta
from pathlib import Path
from urllib.parse import quote
from http.server import HTTPServer, ThreadingHTTPServer, BaseHTTPRequestHandler
import threading
import requests
import openpyxl
from openpyxl import Workbook

# Fix Windows console encoding
if sys.platform == "win32":
    os.system("chcp 65001 >nul 2>&1")
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')

# ═══════════════════════ CONFIG ═══════════════════════
MAXIS_API = "https://maxis-service-prod-dub.amazon.com"
TICKETY_API = "https://midway.us-west-2.cti.api.tickety.amazon.dev/Default/Default"
# Scope the fetch to the exact CTI (GCOI / WFM - Instore / CIPS Create|Update|Subscribe)
# AND the auto-ticket application identity, so other-marketplace / manual / -test tickets
# that merely share a similar TITLE are never picked up.
QUEUE_QUERY = ('extensions.tt.category:"GCOI" AND extensions.tt.type:"WFM - Instore" '
               'AND extensions.tt.item:("CIPS Create" OR "CIPS Update" OR "CIPS Subscribe") '
               'AND extensions.tt.status:(Assigned OR "Work In Progress" OR Researching OR Pending) '
               'AND requesterIdentity:"ticketing-application:wfm-cips-rbs"')
# The auto-ticket application identity used to confirm a ticket is machine-generated.
AUTO_REQUESTER = "ticketing-application:wfm-cips-rbs"
# Map CTI item -> collation stream.
ITEM_TO_STREAM = {"cips create": "Create", "cips update": "Update", "cips subscribe": "Subscribe"}

SCRIPT_DIR = Path(__file__).parent.resolve()
TEMPLATE_PATH = SCRIPT_DIR / "CIPS Template.xlsx"
MASTER_PATH = SCRIPT_DIR / "MASTER Collation File.xlsx"
CACHE_DIR = SCRIPT_DIR / ".cips_cache"
RUN_LOG_PATH = CACHE_DIR / "run_log.json"
EXCLUDE_PATH = SCRIPT_DIR / "Exclude_scancodes.txt"
MIDWAY_COOKIE = Path.home() / ".midway" / "cookie"

PAGE_SIZE = 100
MAX_PAGES = 50
REQUEST_DELAY = 0.3
DOWNLOAD_DELAY = 0.2
MAX_RETRIES = 3
PROXY_REGIONS = ["iad", "dub", "pdx", "nrt"]
MONTH_NAMES = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]

# Internal ticket config
TEMPLATES = {
    "Create": {"cti_item": "CIPS Create"},
    "Update": {"cti_item": "CIPS Update"},
    "Subscribe": {"cti_item": "CIPS Subscribe"},
}
INTERNAL_ASSIGNEE = "jtteja"

RESOLVE_COMMENTS = {
    "Create": "Hello,\nWe ingested all the item_ids into panda01 catalog as requested. Pending items are being tracked separately.",
    "Update": "Hello,\nAs per the request all the items have been updated. Resolving the SIM.",
    "Subscribe": "Hello,\nAs per the request all items have been subscribed. Resolving the SIM.",
}

# ═══════════════════════ HTTP ═══════════════════════
SESSION = None

def get_session():
    if not MIDWAY_COOKIE.exists():
        print(f"❌ Midway cookie not found at: {MIDWAY_COOKIE}")
        print(f"   Run 'mwinit -s' first, then re-run.")
        sys.exit(1)
    session = requests.Session()
    cookie_text = MIDWAY_COOKIE.read_text().strip()
    for line in cookie_text.split("\n"):
        line = line.strip()
        if not line: continue
        if line.startswith("#HttpOnly_"):
            line = line[len("#HttpOnly_"):]
        elif line.startswith("#"):
            continue
        parts = line.split("\t")
        if len(parts) >= 7:
            domain = parts[0]
            path = parts[2]
            name = parts[5]
            value = parts[6]
            session.cookies.set(name, value, domain=domain, path=path)
    return session

def get_sess():
    global SESSION
    if SESSION is None: SESSION = get_session()
    return SESSION

# ── MAXIS (for fetching tickets + downloading CSVs) ──
def maxis_get(path):
    url = f"{MAXIS_API}{path}"
    url += ("&" if "?" in url else "?") + "maxis%3Aheader%3AAmzn-Version=1.0"
    r = get_sess().get(url, timeout=60, allow_redirects=True)
    if r.status_code in (401, 403):
        raise RuntimeError("Auth failed! Run 'mwinit -s' and re-run.")
    data = r.json()
    if data.get("status") == "error" and "uthenti" in data.get("message","").lower():
        raise RuntimeError("Auth expired! Run 'mwinit -s' then re-run.")
    return data

def maxis_get_retry(path, retries=MAX_RETRIES):
    for attempt in range(retries):
        try: return maxis_get(path)
        except RuntimeError as e:
            if "auth" in str(e).lower(): raise
            if attempt == retries - 1: raise
            time.sleep(1 * (attempt + 1))
        except Exception as e:
            if attempt == retries - 1: raise RuntimeError(f"Network error: {e}")
            time.sleep(2 * (attempt + 1))

def download_binary(url):
    for attempt in range(MAX_RETRIES):
        try:
            r = get_sess().get(url, timeout=120, allow_redirects=True)
            if r.status_code == 200 and len(r.content) > 0: return r.content
        except:
            if attempt == MAX_RETRIES - 1: return None
            time.sleep(1)
    return None

# ── TICKETY (for creating/resolving/updating tickets) ──
def tickety_get(path):
    r = get_sess().get(f"{TICKETY_API}{path}", timeout=30, allow_redirects=True)
    if r.status_code in (401, 403):
        raise RuntimeError("Tickety auth failed! Run 'mwinit -s' and re-run.")
    return r

def tickety_post(path, json_data):
    r = get_sess().post(f"{TICKETY_API}{path}", json=json_data,
                        headers={"Content-Type": "application/json"}, timeout=30, allow_redirects=True)
    if r.status_code in (401, 403):
        raise RuntimeError("Tickety auth failed! Run 'mwinit -s' and re-run.")
    return r

# ═══════════════════════ HELPERS ═══════════════════════
def is_file_locked(filepath):
    if not filepath.exists(): return False
    try:
        with open(filepath, "a"): pass
        return False
    except (IOError, PermissionError): return True

def get_sun_sat_week(dt=None):
    """Week number based on Sun-Sat calendar."""
    if dt is None: dt = datetime.now()
    jan1 = datetime(dt.year, 1, 1)
    jan1_dow = (jan1.weekday() + 1) % 7
    day_of_year = (dt - jan1).days
    return (day_of_year + jan1_dow) // 7 + 1

def fmt_date(ds):
    """Format any date value as DD-MM-YYYY."""
    if not ds or str(ds).strip() in ("", "None", "none"): return ""
    if hasattr(ds, 'strftime'): return ds.strftime("%d-%m-%Y")
    s = str(ds).strip()
    if re.match(r"^\d{2}-\d{2}-\d{4}$", s): return s
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})", s)
    if m: return f"{m.group(3)}-{m.group(2)}-{m.group(1)}"
    try:
        d = datetime.fromisoformat(s.replace("Z", "+00:00"))
        return f"{d.day:02d}-{d.month:02d}-{d.year}"
    except: return ""

def get_week_from_date(date_str):
    """Get Sun-Sat week number from DD-MM-YYYY."""
    if not date_str: return ""
    try:
        parts = str(date_str).split("-")
        if len(parts) == 3 and len(parts[2]) == 4:
            d = datetime(int(parts[2]), int(parts[1]), int(parts[0]))
            return str(get_sun_sat_week(d))
    except: pass
    return ""

def detect_file_type(filename):
    fn = (filename or "").upper()
    if "SUBSCRIBE" in fn: return "Subscribe"
    if "CREATE" in fn: return "Create"
    if "UPDATE" in fn: return "Update"
    return None

def extract_vid(issue):
    """Extract V-ticket ID."""
    for a in issue.get("aliases", []):
        aid = a.get("id", "")
        if re.match(r"^V\d+$", aid, re.I): return aid
    return None

def load_run_log():
    if RUN_LOG_PATH.exists():
        try:
            with open(RUN_LOG_PATH, encoding="utf-8") as f:
                return json.load(f)
        except: pass
    return {"runs": [], "internal_tickets": {}, "resolved_tickets": []}

def save_run_log(log):
    with open(RUN_LOG_PATH, "w", encoding="utf-8") as f:
        json.dump(log, f, indent=2)

# ═══════════════════════ TEMPLATE & MASTER ═══════════════════════
def load_template():
    print(f"📄 Loading template: {TEMPLATE_PATH.name}")
    wb = openpyxl.load_workbook(str(TEMPLATE_PATH), read_only=True, data_only=True)
    templates = {}
    for sn in wb.sheetnames:
        sl = sn.lower().strip()
        ft = "Create" if "create" in sl else "Update" if "update" in sl else "Subscribe" if "subscribe" in sl else None
        if not ft: continue
        headers = [str(c.value).strip() if c.value else "" for c in wb[sn][1]]
        while headers and headers[-1] == "": headers.pop()
        templates[ft] = {"sheet_name": sn, "headers": headers, "col_count": len(headers)}
        print(f"   ✅ {ft}: {len(headers)} columns")
    wb.close()
    return templates

def load_master():
    if not MASTER_PATH.exists():
        print(f"❌ Master file not found: {MASTER_PATH}"); sys.exit(1)
    print(f"📂 Loading master: {MASTER_PATH.name}")
    wb = openpyxl.load_workbook(str(MASTER_PATH), read_only=True, data_only=True)
    master = {}
    for sn in wb.sheetnames:
        sl = sn.lower().strip()
        ft = "Create" if "create" in sl else "Update" if "update" in sl else "Subscribe" if "subscribe" in sl else None
        if not ft: continue
        ws = wb[sn]
        headers = [str(c.value).strip() if c.value else "" for c in ws[1]]
        rows = []
        count = 0
        for row in ws.iter_rows(min_row=2, values_only=True):
            r = [fmt_date(v) if hasattr(v, 'strftime') else (str(v).strip() if v is not None else "") for v in row]
            if any(v for v in r): rows.append(r)
            count += 1
            if count % 5000 == 0:
                print(f"\r   📊 {ft}: reading... {count} rows", end="", flush=True)
        master[ft] = {"sheet_name": sn, "headers": headers, "rows": rows}
        print(f"\r   📊 {ft}: {len(rows)} rows, {len(headers)} columns          ")
    wb.close()
    return master

# ═══════════════════════ FETCH & COLLATE ═══════════════════════
def fetch_pending_tickets():
    print(f"\n🔍 Fetching auto CIPS tickets by CTI (GCOI / WFM - Instore / CIPS *)...")
    # Robust, full URL-encoding of the query (the old partial encoder missed chars like '.').
    encoded = quote(QUEUE_QUERY, safe="")
    tickets = []
    start_token = None
    page = 0
    while page < MAX_PAGES:
        url = f"/issues?q={encoded}&rows={PAGE_SIZE}&sort=createDate%20desc"
        if start_token: url += f"&startToken={start_token}"
        resp = maxis_get_retry(url)
        docs = resp.get("documents", [])
        total = resp.get("totalNumberFound", 0)
        for doc in docs:
            uuid = doc.get("id")
            vid = None
            for a in doc.get("aliases", []):
                if re.match(r"^V\d+$", a.get("id",""), re.I): vid = a["id"]
            if not uuid and not vid: continue
            # Derive stream from the CTI item (authoritative) rather than parsing the title.
            ext_tt = doc.get("extensions", {}).get("tt", {})
            item = str(ext_tt.get("item", "")).strip().lower()
            ft = ITEM_TO_STREAM.get(item)
            # Safety anchor: only accept the auto-ticket application identity.
            requester = str(doc.get("requesterIdentity", "")).strip()
            if requester != AUTO_REQUESTER: continue
            if not ft: continue
            title = doc.get("title", "")
            tickets.append({"uuid": uuid, "vid": vid, "type": ft, "title": title[:200], "createDate": doc.get("createDate","")})
        print(f"   Page {page+1}: {len(tickets)} tickets (of ~{total})    ", end="\r")
        start_token = resp.get("startToken")
        if not start_token or len(docs) < PAGE_SIZE: break
        page += 1
        time.sleep(REQUEST_DELAY)
    # Dedupe by ticket id (stream + requester already validated above).
    seen = {}
    for t in tickets: seen[t["uuid"] or t["vid"]] = t
    tickets = list(seen.values())
    # Exclude already-resolved tickets from previous runs
    run_log = load_run_log()
    resolved_set = set(run_log.get("resolved_tickets", []))
    tickets = [t for t in tickets if (t["uuid"] or t["vid"]) not in resolved_set]
    sub = sum(1 for t in tickets if t["type"]=="Subscribe")
    upd = sum(1 for t in tickets if t["type"]=="Update")
    cre = sum(1 for t in tickets if t["type"]=="Create")
    print(f"\n   ✅ {len(tickets)} auto tickets (Create:{cre} Update:{upd} Subscribe:{sub})")
    return tickets

def download_csv(uuid, attachment):
    stack = attachment.get("stack", "prod-iad")
    region = stack.split("-")[-1] if "-" in stack else "iad"
    if region not in PROXY_REGIONS: region = "iad"
    url = f"https://maxis-file-service-prod-{region}.{region}.proxy.amazon.com/issues/{uuid}/attachments/{attachment['id']}"
    data = download_binary(url)
    if not data: return None
    text = data.decode("utf-8-sig", errors="replace").strip()
    if not text: return None
    lines = text.split("\n")
    skip = 0
    for line in lines[:5]:
        s = line.strip().strip('"')
        if not s: skip += 1; continue
        if re.match(r'^version\s*[=:]', s, re.I): skip += 1; continue
        if s.count(",") == 0 and len(s) < 100: skip += 1; continue
        break
    if skip: text = "\n".join(lines[skip:])
    try:
        reader = csv.DictReader(io.StringIO(text))
        return list(reader)
    except: return None

def map_row(csv_row, master_headers, vid, date_str, week_no):
    col_count = len(master_headers)
    out = [""] * col_count
    csv_lower = {k.lower().strip(): k for k in csv_row.keys() if k is not None}
    for i, h in enumerate(master_headers):
        if not h: continue
        hl = h.lower().strip()
        if hl == "sim id": out[i] = vid or ""
        elif hl == "date": out[i] = date_str or ""
        elif hl in ("week", "week no"): out[i] = week_no or ""
        elif h in csv_row: out[i] = str(csv_row[h] or "")
        elif hl in csv_lower: out[i] = str(csv_row[csv_lower[hl]] or "")
    return out

def process_tickets(tickets, master):
    new_data = {"Create": [], "Update": [], "Subscribe": []}
    errors = 0
    t_start = time.time()
    for idx, ticket in enumerate(tickets):
        tid = ticket.get("uuid") or ticket.get("vid")
        if not tid: continue
        elapsed = time.time() - t_start
        eta = f"~{(elapsed/max(idx,1))*(len(tickets)-idx):.0f}s left" if idx > 0 else ""
        print(f"  [{idx+1}/{len(tickets)}] {tid[:24]}... ", end="", flush=True)
        cf = CACHE_DIR / f"{tid}.json"
        if cf.exists():
            # Re-download if cache is older than 7 days
            cache_age = time.time() - cf.stat().st_mtime
            if cache_age < 7 * 86400:
                try:
                    with open(cf, encoding="utf-8") as f:
                        cached = json.load(f)
                    for ft, rows in cached.items():
                        if ft in new_data and rows: new_data[ft].extend(rows)
                    total_r = sum(len(v) for v in cached.values() if isinstance(v, list))
                    print(f"💾 ({total_r} rows) | {eta}")
                    continue
                except: pass
        try:
            issue = maxis_get_retry(f"/issues/{tid}")
            real_uuid = issue.get("id", tid)
            vid = extract_vid(issue) or tid
            date_str = fmt_date(issue.get("createDate"))
            week_no = get_week_from_date(date_str)
            atts = [a for a in issue.get("attachments",[]) if a.get("id") and (a.get("fileName") or "").lower().endswith(".csv")]
            if not atts:
                print(f"📭 no CSV | {eta}")
                with open(cf, "w", encoding="utf-8") as f:
                    json.dump({}, f)
                continue
            ticket_rows = {"Create":[], "Update":[], "Subscribe":[]}
            files_ok = 0; rows_total = 0
            for att in atts:
                ft = detect_file_type(att.get("fileName",""))
                if not ft or ft not in master: continue
                csv_rows = download_csv(real_uuid, att)
                if not csv_rows: errors += 1; continue
                for row in csv_rows:
                    mapped = map_row(row, master[ft]["headers"], vid, date_str, week_no)
                    ticket_rows[ft].append(mapped)
                    rows_total += 1
                files_ok += 1
                time.sleep(DOWNLOAD_DELAY)
            for ft, rows in ticket_rows.items(): new_data[ft].extend(rows)
            with open(cf, "w", encoding="utf-8") as f:
                json.dump(ticket_rows, f)
            print(f"✅ {files_ok} files, {rows_total} rows | {eta}")
        except Exception as e:
            print(f"❌ {str(e)[:45]}")
            errors += 1
        time.sleep(REQUEST_DELAY)
    if errors: print(f"\n   ⚠️ {errors} errors")
    return new_data

# ═══════════════════════ MERGE & DEDUP ═══════════════════════
def merge_and_dedup(master, new_data):
    print("\n📦 Merging and deduplicating (by scancode)...")
    stats = {}
    for ft in ["Create", "Update", "Subscribe"]:
        if ft not in master: continue
        headers = master[ft]["headers"]
        existing_rows = master[ft]["rows"]
        new_rows = new_data.get(ft, [])
        # Dedup key = scancode (external_product_information#2.value), which is unique per item.
        # (Previously deduped on vendor_sku, which wrongly dropped legitimately different items
        #  that a single vendor sells under the same vendor_sku but with different scancodes.)
        sc_idx = next((i for i, h in enumerate(headers) if h.lower().strip() == "external_product_information#2.value"), None)
        if sc_idx is None:
            master[ft]["rows"] = new_rows + existing_rows
            stats[ft] = {"raw": len(new_rows), "unique_new": len(new_rows), "after": len(master[ft]["rows"]), "removed": 0}
            continue
        all_rows = new_rows + existing_rows
        total_before = len(all_rows)
        print(f"   {ft}: deduplicating {total_before} rows by scancode...", end=" ", flush=True)
        seen_scancodes = set()
        deduped_rows = []
        new_unique = 0
        for i, row in enumerate(all_rows):
            scancode = row[sc_idx].strip().lower() if sc_idx < len(row) and row[sc_idx] else ""
            if not scancode:
                deduped_rows.append(row)
                if i < len(new_rows): new_unique += 1
            elif scancode not in seen_scancodes:
                seen_scancodes.add(scancode)
                deduped_rows.append(row)
                if i < len(new_rows): new_unique += 1
        removed = total_before - len(deduped_rows)
        master[ft]["rows"] = deduped_rows
        stats[ft] = {"raw": len(new_rows), "unique_new": new_unique, "after": len(deduped_rows), "removed": removed}
        print(f"✅")
        print(f"   {ft}: {len(existing_rows)} + {len(new_rows)} new → {len(deduped_rows)} unique (-{removed})")
    return stats

def save_master(master):
    print(f"\n📝 Saving {MASTER_PATH.name}...")
    backup = MASTER_PATH.with_suffix(".xlsx.bak")
    if MASTER_PATH.exists():
        shutil.copy2(MASTER_PATH, backup)
        print(f"   💾 Backup: {backup.name}")
    wb = Workbook(write_only=True)
    for ft in ["Create", "Subscribe", "Update"]:
        if ft not in master: continue
        info = master[ft]
        total = len(info['rows'])
        print(f"   Writing {ft} ({total} rows)...", end="", flush=True)
        ws = wb.create_sheet(title=info["sheet_name"])
        ws.append(info["headers"])
        col_count = len(info["headers"])
        for i, row in enumerate(info["rows"]):
            if len(row) < col_count: row = row + [""] * (col_count - len(row))
            elif len(row) > col_count: row = row[:col_count]
            ws.append(row)
            if (i+1) % 10000 == 0: print(f"\r   Writing {ft} ({total} rows)... {i+1}/{total}", end="", flush=True)
        print(f"\r   Writing {ft} ({total} rows)... ✅                    ")
    print(f"   Saving...", end=" ", flush=True)
    wb.save(str(MASTER_PATH)); wb.close()
    print(f"✅ ({MASTER_PATH.stat().st_size/(1024*1024):.1f} MB)")

# ═══════════════════════ TICKET OPERATIONS (Tickety API) ═══════════════════════
def maxis_edit(ticket_uuid, path_edits):
    """Apply edits to a ticket via MAXIS edits endpoint (for assignee, severity, etc)."""
    try:
        url = f"{MAXIS_API}/issues/{ticket_uuid}/edits"
        r = get_sess().post(url, json={"pathEdits": path_edits},
                           headers={"Content-Type": "application/json", "Accept": "application/json"},
                           timeout=15, allow_redirects=True)
        return r.status_code in (200, 201)
    except:
        return False

def create_internal_ticket(stream, raw, final, week_no):
    """Create an internal tracking ticket. Returns ticket ID or None."""
    tpl = TEMPLATES[stream]
    title = f"WFM CIPS Collation - {stream}  | Week {week_no} (for internal GCOI tracking only)"
    description = f"(for internal GCOI tracking only)\n\nTotal raw sku count: {raw}\nTotal final sku count: {final}"
    
    # Use MAXIS REST to create (supports extensions.tt with custom fields + avoids AutoSIM bot)
    payload = {
        "title": title,
        "description": description,
        "descriptionContentType": "text/plain",
        "assignedFolder": "c9f6dd71-2bf2-4b64-8725-9dfa3de99a35",
        "extensions": {
            "tt": {
                "category": "GCOI",
                "type": "WFM - Instore",
                "item": tpl["cti_item"],
                "assignedGroup": "gcoi-wfm",
                "caseType": "Trouble Ticket",
                "building": "HYD13-SEZCO(Hyderabad, TG, IN)",
                "buildingId": "HYD13",
                "city": "Hyderabad",
                "country": "India",
                "customFields": {
                    "number": [
                        {"id": "raw_sku_count", "value": raw},
                        {"id": "final_sku_count", "value": final},
                    ]
                }
            }
        }
    }
    try:
        url = f"{MAXIS_API}/issues?maxis%3Aheader%3AAmzn-Version=1.0"
        r = get_sess().post(url, json=payload, headers={"Content-Type": "application/json", "Accept": "application/json"}, timeout=30, allow_redirects=True)
        if r.status_code in (200, 201):
            data = r.json()
            ticket_uuid = data.get("id", "")
            ticket_id = ticket_uuid
            
            # Small delay to let MAXIS index the ticket
            time.sleep(1)
            
            # Get the real UUID by fetching the ticket
            try:
                r_fetch = get_sess().get(f"{MAXIS_API}/issues/{ticket_uuid}?maxis%3Aheader%3AAmzn-Version=1.0",
                                        headers={"Accept": "application/json"}, timeout=15, allow_redirects=True)
                if r_fetch.status_code == 200:
                    fetch_data = r_fetch.json()
                    real_uuid = fetch_data.get("id", ticket_uuid)
                    # Get D-alias
                    for a in fetch_data.get("aliases", []):
                        if re.match(r"^D\d+$", a.get("id", "")): ticket_id = a["id"]; break
                else:
                    real_uuid = ticket_uuid
            except:
                real_uuid = ticket_uuid
            
            # Set severity (SEV-5) and assignee via edits endpoint
            maxis_edit(real_uuid, [
                {"path": "/extensions/tt/impact", "editAction": "PUT", "data": 5},
                {"path": "/autoUpgrade/workingHours", "editAction": "PUT", "data": False},
            ])
            maxis_edit(real_uuid, [
                {"path": "/assigneeIdentity", "editAction": "PUT", "data": f"kerberos:{INTERNAL_ASSIGNEE}@ANT.AMAZON.COM"},
            ])
            
            # Fetch D-alias
            try:
                r2 = get_sess().get(f"{MAXIS_API}/issues/{ticket_uuid}?maxis%3Aheader%3AAmzn-Version=1.0",
                                    headers={"Accept": "application/json"}, timeout=15, allow_redirects=True)
                if r2.status_code == 200:
                    for a in r2.json().get("aliases", []):
                        if re.match(r"^D\d+$", a.get("id", "")): ticket_id = a["id"]; break
            except: pass
            return ticket_id
        print(f"      ❌ Create failed: {r.status_code} {r.text[:100]}")
    except Exception as e:
        print(f"      ❌ Error: {str(e)[:50]}")
    return None

def update_internal_ticket(ticket_id, raw, final):
    """Update an existing internal ticket's description and custom fields."""
    description = f"(for internal GCOI tracking only)\n\nTotal raw sku count: {raw}\nTotal final sku count: {final}"
    r = tickety_post(f"/tickets/{ticket_id}", {"description": description})
    if r.status_code != 200:
        return False
    
    # Update custom fields via GraphQL (needs UUID, not D-ID)
    # Fetch UUID from ticket
    try:
        issue = maxis_get_retry(f"/issues/{ticket_id}")
        uuid = issue.get("id", "")
        if uuid:
            update_custom_fields(uuid, raw, final)
    except: pass
    return True

def update_custom_fields(uuid, raw, final):
    """Update raw_sku_count and final_sku_count custom fields via GraphQL."""
    GRAPHQL_URL = "https://sim-ticketing-graphql-fleet.corp.amazon.com/graphql"
    payload = [{
        "operationName": "updateCustomFields",
        "variables": {
            "args": {
                "ticketId": uuid,
                "customFields": [
                    {"id": "raw_sku_count", "value": {"integerValue": raw}},
                    {"id": "final_sku_count", "value": {"integerValue": final}}
                ]
            }
        },
        "query": "mutation updateCustomFields($args: UpdateCustomFieldsInputTickety!) {\n  updateCustomFieldsTickety(args: $args) {\n    id\n    value {\n      integerValue\n      __typename\n    }\n    __typename\n  }\n}\n"
    }]
    try:
        r = get_sess().post(GRAPHQL_URL, json=payload,
            headers={"Content-Type": "application/json", "Accept": "application/json"},
            timeout=30, allow_redirects=True)
        return r.status_code == 200
    except:
        return False

def resolve_ticket(ticket_id, comment):
    """Add comment and resolve a ticket."""
    try:
        # Add comment via Tickety API
        r1 = tickety_post(f"/tickets/{ticket_id}/comments", {
            "message": comment, "threadName": "CORRESPONDENCE", "contentType": "text/plain"
        })
        if r1.status_code not in (200, 201):
            print(f"\n      Comment failed ({ticket_id}): {r1.status_code} {r1.text[:80]}")
            return False
        
        # Resolve via Tickety API with correct field name
        r2 = tickety_post(f"/tickets/{ticket_id}", {
            "status": "Resolved",
            "logTimeSpentInMinutes": 30,
            "rootCause": "N/A"
        })
        if r2.status_code == 200:
            return True
        
        print(f"\n      Resolve failed ({ticket_id}): {r2.status_code} {r2.text[:80]}")
        return False
    except Exception as e:
        print(f"\n      Error ({ticket_id}): {str(e)[:50]}")
        return False

def find_existing_internal_ticket(stream, week_no):
    """Search for an existing internal collation ticket for this stream/week (only open/assigned ones)."""
    # Simple search without status filter (we'll filter results ourselves)
    query = f'title:"WFM CIPS Collation - {stream}" AND title:"Week {week_no}" AND extensions.tt.assignedGroup:"gcoi-wfm"'
    encoded = query.replace(" ","%20").replace('"','%22')
    try:
        resp = maxis_get_retry(f"/issues?q={encoded}&rows=5&sort=createDate%20desc")
        docs = resp.get("documents", [])
        for doc in docs:
            title = doc.get("title", "")
            status = doc.get("status", "") or (doc.get("extensions", {}).get("tt", {}).get("status", ""))
            # Only match open/assigned tickets, skip resolved/closed, skip TEST tickets
            if f"Week {week_no}" in title and "Collation" in title and "TEST" not in title.upper():
                if status in ("Resolved", "Closed"):
                    continue
                ticket_id = doc.get("id", "")
                vid = None
                for a in doc.get("aliases", []):
                    if re.match(r"^D\d+$", a.get("id","")): vid = a["id"]
                return vid or ticket_id
    except Exception as e:
        print(f"      Search error: {str(e)[:50]}")
    return None

def read_internal_ticket_counts(ticket_id):
    """Read the current (raw, final) sku counts from an existing internal ticket.

    Tries the custom fields (raw_sku_count / final_sku_count) first, then falls back to
    parsing the description ('Total raw sku count: N' / 'Total final sku count: N').
    Returns (raw, final) as ints; either may be None if it could not be read.
    """
    raw = final = None
    try:
        issue = maxis_get_retry(f"/issues/{ticket_id}")
    except Exception as e:
        print(f"      Could not read counts for {ticket_id}: {str(e)[:50]}")
        return (None, None)

    # 1) Custom fields (always populated on these tickets).
    try:
        cfs = issue.get("extensions", {}).get("tt", {}).get("customFields", {})
        # customFields may be {"number":[{"id":...,"value":...}, ...]} or a flat list.
        entries = []
        if isinstance(cfs, dict):
            for v in cfs.values():
                if isinstance(v, list): entries.extend(v)
        elif isinstance(cfs, list):
            entries = cfs
        for e in entries:
            if not isinstance(e, dict): continue
            cid = e.get("id", "")
            val = e.get("value")
            if val is None: continue
            try: ival = int(str(val).strip())
            except: continue
            if cid == "raw_sku_count": raw = ival
            elif cid == "final_sku_count": final = ival
    except: pass

    # 2) Description fallback for whichever value is still missing.
    if raw is None or final is None:
        try:
            desc = issue.get("description", "") or ""
            m_raw = re.search(r"raw sku count\s*[:=]\s*(\d+)", desc, re.I)
            m_fin = re.search(r"final sku count\s*[:=]\s*(\d+)", desc, re.I)
            if raw is None and m_raw: raw = int(m_raw.group(1))
            if final is None and m_fin: final = int(m_fin.group(1))
        except: pass

    return (raw, final)

def handle_internal_tickets(stats, new_data, tickets_processed, run_log, master=None, imp_results=None):
    """Raise/update internal tickets based on actual weeks in collated data.
    Final = Raw - Backlog (backlog = pending rows from Pending CIPS sheet for that week)."""
    print(f"\n🎫 Internal tickets...")
    
    # Group collated tickets by week number (from their creation date)
    week_stats = defaultdict(lambda: {"Create": {"raw": 0}, "Update": {"raw": 0}, "Subscribe": {"raw": 0}})
    
    for ticket in tickets_processed:
        ft = ticket.get("type")
        if not ft: continue
        cd = ticket.get("createDate", "")
        date_str = fmt_date(cd)
        wk = get_week_from_date(date_str)
        if not wk: continue
        # Count raw rows from this ticket (from cache)
        tid = ticket.get("uuid") or ticket.get("vid")
        cf = CACHE_DIR / f"{tid}.json"
        if cf.exists():
            try:
                with open(cf, encoding="utf-8") as f:
                    cached = json.load(f)
                rows = cached.get(ft, [])
                week_stats[int(wk)][ft]["raw"] += len(rows)
            except: pass
    
    # Count backlog per week per stream from Pending CIPS sheet
    # The pending sheet has rows from master that are NOT in IMP
    # We need to count how many of THOSE belong to each week
    week_backlog = defaultdict(lambda: {"Create": 0, "Update": 0, "Subscribe": 0})
    if master and imp_results:
        for ft in ["Create", "Update", "Subscribe"]:
            if ft not in master: continue
            headers = master[ft]["headers"]
            epi2_idx = next((i for i, h in enumerate(headers) if h.lower().strip() == "external_product_information#2.value"), None)
            week_idx = next((i for i, h in enumerate(headers) if h.lower().strip() in ("week", "week no")), None)
            if week_idx is None: continue
            
            for row in new_data.get(ft, []):
                # A row is pending (backlog) unless its scancode is DONE per the bridge status.
                is_pending = False
                if epi2_idx is None:
                    is_pending = True
                else:
                    sc = row[epi2_idx].strip() if epi2_idx < len(row) and row[epi2_idx] else ""
                    if not sc or sc == "None":
                        is_pending = True
                    else:
                        st = imp_results.get(sc)
                        if not isinstance(st, dict) or st.get("status") != "done":
                            is_pending = True
                
                if is_pending:
                    wk = row[week_idx].strip() if week_idx < len(row) and row[week_idx] else ""
                    if wk:
                        try: week_backlog[int(wk)][ft] += 1
                        except: pass
    
    weeks_found = sorted(week_stats.keys())
    if not weeks_found:
        print("   No week data found in collated tickets.")
        return
    
    print(f"   Weeks in data: {weeks_found}")
    if imp_results:
        for ft in ["Create", "Update", "Subscribe"]:
            raw_t = stats.get(ft, {}).get("raw", 0)
            if raw_t > 0:
                backlog_t = sum(week_backlog[w][ft] for w in weeks_found)
                print(f"   {ft}: raw={raw_t}, final={raw_t - backlog_t}, backlog={backlog_t}")
    year = datetime.now().year
    
    if "internal_tickets" not in run_log: run_log["internal_tickets"] = {}
    
    for wk in weeks_found:
        week_key = f"{year}-W{wk}"
        existing = run_log.get("internal_tickets", {}).get(week_key, {})
        
        for ft in ["Create", "Update", "Subscribe"]:
            raw = week_stats[wk][ft]["raw"]
            if raw == 0: continue
            
            # Final = Raw - Backlog for this week
            backlog = week_backlog[wk][ft]
            final = raw - backlog
            
            prev = existing.get(ft, {})
            
            if prev and prev.get("ticket_id"):
                # Update existing from cache
                ticket_id = prev["ticket_id"]
                cumulative_raw = prev.get("raw", 0) + raw
                cumulative_final = prev.get("final", 0) + final
                print(f"   Week {wk} {ft}: Updating {ticket_id} (raw +{raw}={cumulative_raw})")
                if update_internal_ticket(ticket_id, cumulative_raw, cumulative_final):
                    existing[ft] = {"ticket_id": ticket_id, "raw": cumulative_raw, "final": cumulative_final}
                    print(f"   ✅ Updated")
                else:
                    print(f"   ❌ Update failed")
            else:
                # Search for existing ticket
                found_id = find_existing_internal_ticket(ft, wk)
                if found_id:
                    # Read the ticket's CURRENT counts and APPEND this run's numbers, so a
                    # mid-week re-run adds to the existing total instead of overwriting it.
                    prev_raw, prev_final = read_internal_ticket_counts(found_id)
                    base_raw = prev_raw if prev_raw is not None else 0
                    base_final = prev_final if prev_final is not None else 0
                    cumulative_raw = base_raw + raw
                    cumulative_final = base_final + final
                    print(f"   Week {wk} {ft}: Found {found_id} — appending (raw {base_raw}+{raw}={cumulative_raw}, final {base_final}+{final}={cumulative_final})")
                    if update_internal_ticket(found_id, cumulative_raw, cumulative_final):
                        existing[ft] = {"ticket_id": found_id, "raw": cumulative_raw, "final": cumulative_final}
                        print(f"   ✅ Updated existing")
                    else:
                        print(f"   ❌ Update failed")
                else:
                    # Create new
                    print(f"   Week {wk} {ft}: Creating (raw={raw}, final={final})")
                    ticket_id = create_internal_ticket(ft, raw, final, wk)
                    if ticket_id:
                        existing[ft] = {"ticket_id": ticket_id, "raw": raw, "final": final}
                        print(f"   ✅ Created {ticket_id}")
                    else:
                        print(f"   ❌ Create failed")
            
            time.sleep(1)
        
        run_log["internal_tickets"][week_key] = existing

def resolve_source_tickets(tickets, run_log):
    """Resolve all collated source tickets."""
    print(f"\n✅ Resolving {len(tickets)} source tickets...")
    resolved_set = set(run_log.get("resolved_tickets", []))
    to_resolve = [t for t in tickets if (t["uuid"] or t["vid"]) not in resolved_set]
    
    if not to_resolve:
        print(f"   All already resolved!")
        return 0
    
    print(f"   {len(to_resolve)} to resolve ({len(tickets) - len(to_resolve)} already done)")
    
    # Confirmation prompt
    confirm = input(f"\n   Resolve {len(to_resolve)} tickets? (Y/n): ").strip().lower()
    if confirm not in ("y", "yes", ""):
        print("   ⏭️  Skipped resolving.")
        return 0
    
    ok = 0; fail = 0
    for idx, ticket in enumerate(to_resolve):
        # Prefer V-ticket ID for Tickety API (more reliable)
        tid = ticket.get("vid") or ticket.get("uuid")
        ft = ticket.get("type")
        if not tid or not ft: continue
        comment = RESOLVE_COMMENTS.get(ft, "Resolving the SIM.")
        if resolve_ticket(tid, comment):
            ok += 1
            resolved_set.add(ticket.get("uuid") or ticket.get("vid"))
        else:
            fail += 1
        if (idx+1) % 10 == 0:
            print(f"   [{idx+1}/{len(to_resolve)}] {ok} ok, {fail} failed", end="\r")
        time.sleep(REQUEST_DELAY)
    
    run_log["resolved_tickets"] = list(resolved_set)[-2000:]  # Keep last 2000 to prevent unbounded growth
    print(f"\n   ✅ Resolved: {ok}/{len(to_resolve)}" + (f" (⚠️ {fail} failed)" if fail else ""))
    return ok

# ═══════════════════════ IMP CHECK (Localhost Bridge v4 — single-tab grocerycentral) ═══════════════════════
# Bridge protocol (matches "CIPS IMP Bridge" v4.0 userscript):
#   GET  /status     -> {"status":"ready"}            (bridge probes until server is up)
#   GET  /scancodes  -> {"scancodes":[...]}           (bridge pulls the job)
#   POST /imp-chunk  <- {"results":{sc:{asin,ggSku}}, "processed":N, "total":M}
#   POST /imp-done   <- {"total":M}                   (bridge signals completion)
# Python applies the done/pending rules itself from {asin, ggSku}.
IMP_STATE = {"scancodes": [], "done": False, "processed": 0, "total": 0, "results": {},
             "last_chunk_time": 0, "connected": False, "scancodes_fetched": False}

class IMPBridgeHandler(BaseHTTPRequestHandler):
    def log_message(self, format, *args): pass  # Suppress logs

    def _mark_connected(self):
        if not IMP_STATE.get("connected"):
            IMP_STATE["connected"] = True
            print("   Bridge connected.")

    def _respond(self, code, data=None):
        self.send_response(code)
        self.send_header('Content-Type', 'application/json')
        self.send_header('Access-Control-Allow-Origin', '*')
        self.end_headers()
        if data: self.wfile.write(json.dumps(data).encode())

    def do_GET(self):
        if self.path == '/status':
            self._mark_connected()
            self._respond(200, {"status": "ready"})
        elif self.path == '/scancodes':
            self._mark_connected()
            if not IMP_STATE.get("scancodes_fetched"):
                IMP_STATE["scancodes_fetched"] = True
                print(f"   Bridge picked up {len(IMP_STATE.get('scancodes', []))} scancodes.")
            self._respond(200, {"scancodes": IMP_STATE.get("scancodes", [])})
        elif self.path == '/progress':
            self._respond(200, {"processed": IMP_STATE["processed"], "total": IMP_STATE["total"], "done": IMP_STATE["done"]})
        else:
            self.send_response(404)
            self.end_headers()

    def do_POST(self):
        content_length = int(self.headers.get('Content-Length', 0))
        body = self.rfile.read(content_length)

        if self.path == '/imp-chunk':
            self._mark_connected()
            try:
                chunk = json.loads(body)
                results = chunk.get("results", {})
                IMP_STATE["results"].update(results)
                IMP_STATE["processed"] = chunk.get("processed", IMP_STATE["processed"])
                IMP_STATE["total"] = chunk.get("total", IMP_STATE["total"])
                IMP_STATE["last_chunk_time"] = time.time()
            except: pass
            self._respond(200, {"status": "ok"})
        elif self.path == '/imp-done':
            IMP_STATE["done"] = True
            IMP_STATE["last_chunk_time"] = time.time()
            self._respond(200, {"status": "ok"})
        else:
            self.send_response(404)
            self.end_headers()

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()

def run_imp_check(master):
    """Run the IMP-only check via the single-tab grocerycentral browser bridge (v4).

    The bridge returns, per scancode, a small object: { "asin": str, "ggSku": str, "inactive": bool }.
    Python applies the done/pending rules here:

        DONE    if the scancode is INACTIVE in IMP, OR
                if the scancode has its own ASIN in IMP, OR
                if it is a child (has a globalGroupingSku / parent scancode) whose
                parent scancode differs from the scancode itself.
        PENDING otherwise (active, no ASIN, and not a child of a different parent), or if the
                bridge returned no data for it (safe default — keep it in the backlog).

    Returns dict of scancode -> status object:
        { "status": "done"|"pending", "asin": str, "ggSku": str, "inactive": bool }

    NO CACHING: every run re-checks 100% of master scancodes, because a previously-mapped
    ASIN can silently vanish (associate/team error) and must be re-detected each run.
    """
    print(f"\nIMP check")

    # Collect ALL scancodes from master (external_product_information#2.value). No cache/skip.
    all_scancodes = set()
    for ft in ["Create", "Update", "Subscribe"]:
        if ft not in master: continue
        headers = master[ft]["headers"]
        epi2_idx = next((i for i, h in enumerate(headers) if h.lower().strip() == "external_product_information#2.value"), None)
        if epi2_idx is None: continue
        for row in master[ft]["rows"]:
            if epi2_idx < len(row) and row[epi2_idx]:
                sc = str(row[epi2_idx]).strip()
                if sc and sc != "None":
                    all_scancodes.add(sc)

    to_check = sorted(all_scancodes)
    print(f"   Scancodes to check: {len(to_check)}")
    if not to_check:
        print("   No scancodes in master to check.")
        return {}

    # Reset bridge state
    IMP_STATE["scancodes"] = to_check
    IMP_STATE["done"] = False
    IMP_STATE["processed"] = 0
    IMP_STATE["total"] = len(to_check)
    IMP_STATE["results"] = {}
    IMP_STATE["last_chunk_time"] = time.time()
    IMP_STATE["connected"] = False
    IMP_STATE["scancodes_fetched"] = False

    # Start localhost server FIRST, then open the browser so the bridge finds it immediately.
    # Threaded so incoming chunk/heartbeat POSTs are handled promptly (not queued behind others).
    server = ThreadingHTTPServer(('localhost', 5050), IMPBridgeHandler)
    server_thread = threading.Thread(target=server.serve_forever, daemon=True)
    server_thread.start()

    # Open a fresh grocerycentral tab. The v4 bridge userscript runs there (single tab,
    # same-origin IMP fetch — the fast path), pulls scancodes, and streams {asin, ggSku}
    # chunks back. Requires: bridge script installed, grocerycentral logged in, localhost
    # allowed in Tampermonkey.
    webbrowser.open('https://grocerycentral.amazon.dev/')
    print("   Waiting for browser bridge (opening Grocery Central tab)...")

    # Simple idle-timeout wait: keep going while the bridge is sending chunks/heartbeats; give up
    # only after a long silence. The v4 bridge heartbeats every ~10k scancodes during the IMP
    # fetch, so this should never fire mid-run — 600s is a safety cushion for slow IMP retries.
    idle_timeout = 600
    last_line = ""

    while not IMP_STATE["done"]:
        time.sleep(1)
        now = time.time()
        if now - IMP_STATE["last_chunk_time"] > idle_timeout:
            print(f"\n   Bridge silent for {int(now - IMP_STATE['last_chunk_time'])}s — saving partial results "
                  f"({IMP_STATE['processed']}/{len(to_check)}).")
            break
        if IMP_STATE.get("connected"):
            d, t = IMP_STATE["processed"], len(to_check)
            pct = (d * 100 // t) if t else 0
            got = len(IMP_STATE["results"])
            if d >= t and got < t:
                # Heartbeats have reached 100%; now the real result chunks are streaming in.
                line = f"   [receiving results] {got}/{t}"
            else:
                line = f"   [running] {d}/{t} ({pct}%)"
            if line != last_line:
                print("\r" + line + " " * 16, end="", flush=True)
                last_line = line
    print()  # newline after the in-place status line

    server.shutdown()

    results = IMP_STATE["results"]
    if not results:
        print("   No results received from browser.")
        return {}

    # Apply the done/pending rules. Anything the bridge didn't return -> pending (safe default).
    normalized = {}
    done_n = 0
    for sc in all_scancodes:
        r = results.get(sc)
        if isinstance(r, dict):
            asin = str(r.get("asin", "") or "").strip()
            gg = str(r.get("ggSku", "") or "").strip()
            inactive = r.get("inactive") is True
            # Rule 1: inactive in IMP -> done.
            # Rule 2: has its own ASIN -> done.
            # Rule 3: is a child (has a parent scancode) whose parent != itself -> done.
            is_done = inactive or bool(asin) or (bool(gg) and gg != sc)
            normalized[sc] = {"status": "done" if is_done else "pending", "asin": asin, "ggSku": gg, "inactive": inactive}
            if is_done: done_n += 1
        else:
            normalized[sc] = {"status": "pending", "asin": "", "ggSku": "", "inactive": False}

    pending_n = len(normalized) - done_n
    print(f"   Check complete: {done_n} done, {pending_n} pending")
    return normalized

def load_exclude_scancodes():
    """Read Exclude_scancodes.txt from the CIPS folder.

    Format: one scancode per line; commas/spaces/tabs also accepted as separators.
    Blank lines and non-numeric tokens are ignored. Returns a set of scancode strings
    (empty set if the file is missing or has no valid scancodes).
    """
    if not EXCLUDE_PATH.exists():
        return set()
    try:
        text = EXCLUDE_PATH.read_text(encoding="utf-8-sig", errors="replace")
    except Exception as e:
        print(f"   ⚠️ Could not read {EXCLUDE_PATH.name}: {str(e)[:60]}")
        return set()
    tokens = re.split(r"[\s,;|]+", text)
    return {t.strip() for t in tokens if t.strip().isdigit()}

def generate_pending_sheet(master, imp_results, exclude_set=None):
    """Generate Pending CIPS.xlsx: pure pending rows per stream (original master layout).

    A scancode is DONE (excluded from pending) only when imp_results[sc]['status'] == 'done'
    (has its own ASIN in IMP, OR is a child whose parent scancode differs from itself).
    Everything else stays PENDING.

    exclude_set (optional): a set of scancodes to drop from the Pending sheet REGARDLESS of
    their IMP status. These rows are still counted as backlog for internal tickets (handled
    separately in handle_internal_tickets); this only keeps them out of the handoff sheet.
    """
    print(f"\nGenerating Pending sheet")
    exclude_set = exclude_set or set()

    def status_of(sc):
        r = imp_results.get(sc)
        return r if isinstance(r, dict) else None

    pending_path = SCRIPT_DIR / "Pending CIPS.xlsx"
    wb = Workbook(write_only=True)
    total_pending = 0
    total_excluded = 0

    for ft in ["Create", "Update", "Subscribe"]:
        if ft not in master: continue
        headers = master[ft]["headers"]
        epi2_idx = next((i for i, h in enumerate(headers) if h.lower().strip() == "external_product_information#2.value"), None)

        pending_rows = []
        for row in master[ft]["rows"]:
            if epi2_idx is None:
                pending_rows.append(row)  # can't identify scancode → keep as pending
                continue
            sc = str(row[epi2_idx]).strip() if epi2_idx < len(row) else ""
            # Exclusion: drop from the handoff sheet regardless of IMP status.
            if sc and sc in exclude_set:
                total_excluded += 1
                continue
            if not sc or sc == "None":
                pending_rows.append(row)  # no scancode → pending
                continue
            st = status_of(sc)
            if st is None:
                pending_rows.append(row)  # not checked → pending (safe default)
                continue
            if st.get("status") != "done":
                pending_rows.append(row)
            # else status == done → excluded from pending

        if pending_rows:
            ws = wb.create_sheet(title=master[ft]["sheet_name"])
            ws.append(headers)
            col_count = len(headers)
            for row in pending_rows:
                if len(row) < col_count: row = row + [""] * (col_count - len(row))
                elif len(row) > col_count: row = row[:col_count]
                ws.append(row)
            total_pending += len(pending_rows)
            print(f"   {ft}: {len(pending_rows)} pending rows")

    wb.save(str(pending_path))
    wb.close()
    if exclude_set:
        print(f"   Excluded {total_excluded} rows (from Exclude_scancodes.txt)")
    print(f"   Saved: {pending_path.name} ({total_pending} pending rows)")
    return total_pending

# ═══════════════════════ MAIN ═══════════════════════
def main():
    print("╔═══════════════════════════════════════════════════════╗")
    print("║   CIPS Auto Collator v2.7 — Full Automation           ║")
    print("╚═══════════════════════════════════════════════════════╝")
    today = datetime.now()
    week_no = get_sun_sat_week(today)
    print(f"   Date: {today.strftime('%d-%m-%Y')} | Week {week_no}")
    print(f"   Template: {TEMPLATE_PATH.name}")
    print(f"   Master: {MASTER_PATH.name}")
    print()

    # Pre-checks
    if not TEMPLATE_PATH.exists():
        print(f"❌ Template not found: {TEMPLATE_PATH}"); input("\nPress Enter..."); sys.exit(1)
    if not MASTER_PATH.exists():
        print(f"❌ Master file not found: {MASTER_PATH}"); input("\nPress Enter..."); sys.exit(1)
    if is_file_locked(MASTER_PATH):
        print(f"⚠️  Master file is open! Close it and re-run."); input("\nPress Enter..."); sys.exit(1)
    
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    templates = load_template()
    master = load_master()
    run_log = load_run_log()

    # Test auth
    print("\n🔐 Testing auth...")
    try:
        r = tickety_get("/tickets/V2123197916/summary")
        assert r.status_code == 200
        print("   ✅ Tickety OK")
    except Exception as e:
        print(f"   ❌ {e}\n   Run 'mwinit -s' then re-run."); input("\nPress Enter..."); sys.exit(1)

    # Fetch
    tickets = fetch_pending_tickets()
    if not tickets:
        print("\n📭 No pending tickets. Nothing to do.")
        input("\nPress Enter..."); return

    # Process
    print(f"\n{'═'*55}\n📦 Processing {len(tickets)} tickets\n{'═'*55}")
    t0 = time.time()
    new_data = process_tickets(tickets, master)
    print(f"\n   ⏱️ {time.time()-t0:.0f}s")

    new_total = sum(len(v) for v in new_data.values())
    if new_total == 0:
        print("\n📭 No CSV data found.")
        input("\nPress Enter..."); return
    print(f"\n   📊 New: Create:{len(new_data['Create'])} Update:{len(new_data['Update'])} Subscribe:{len(new_data['Subscribe'])}")

    # Merge & Dedup
    stats = merge_and_dedup(master, new_data)
    save_master(master)

    # IMP Check (before resolve — need results for internal ticket final numbers)
    imp_results = run_imp_check(master)

    # Exclude scancodes (Pending sheet only — does NOT affect internal ticket numbers)
    exclude_set = set()
    confirm_excl = input("\n   Do you want to exclude scancodes from the Pending sheet? (Y/n): ").strip().lower()
    if confirm_excl in ("y", "yes", ""):
        exclude_set = load_exclude_scancodes()
        if exclude_set:
            print(f"   Loaded {len(exclude_set)} scancodes from {EXCLUDE_PATH.name}")
        else:
            print(f"   No scancodes found in {EXCLUDE_PATH.name} — nothing to exclude.")
    else:
        print("   ⏭️  Skipping exclusion.")

    pending_count = 0
    if imp_results:
        pending_count = generate_pending_sheet(master, imp_results, exclude_set)

    # Internal tickets (with IMP-aware backlog numbers)
    handle_internal_tickets(stats, new_data, tickets, run_log, master, imp_results)

    # Resolve source tickets (last step)
    resolve_source_tickets(tickets, run_log)

    # Save run log
    run_log["runs"].append({
        "date": today.isoformat(),
        "week": week_no,
        "tickets_processed": len(tickets),
        "stats": stats,
        "pending_count": pending_count
    })
    save_run_log(run_log)

    # Summary
    elapsed = (time.time() - t0) / 60
    print(f"\n{'═'*55}")
    print(f"🎉 DONE in {elapsed:.1f} min")
    print(f"{'═'*55}")
    print(f"\n📊 Summary:")
    for ft, s in stats.items():
        print(f"   {ft}: {s['raw']} raw, {s['unique_new']} unique new, {s['removed']} dupes removed")
    print(f"\n   Master: {sum(len(master[ft]['rows']) for ft in master)} total rows")
    if imp_results:
        imp_confirmed = sum(1 for v in imp_results.values() if isinstance(v, dict) and v.get("status") == "done")
        imp_pending = sum(1 for v in imp_results.values() if not (isinstance(v, dict) and v.get("status") == "done"))
        print(f"   IMP: {imp_confirmed} done, {imp_pending} pending (backlog)")

    input("\nPress Enter to exit...")

if __name__ == "__main__":
    main()
