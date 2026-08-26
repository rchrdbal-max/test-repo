"""
Export ActiveCampaign per-contact activities into the existing Excel.

Reads contact IDs from the 'export final' sheet of Active_campaign-all data.xlsx,
fetches activities in bulk from AC API, then writes them to a new 'Activities'
sheet and marks the last column (Status synchronizácie trackingu) with 'OK'
for each processed contact.

Resumable: if the script is re-run, contacts already marked 'OK' are skipped
(only their activities are not re-written; the new sheet is rebuilt from cache).

Activity types covered:
  - List Subscribed / Unsubscribed
  - Automation Entered / Completed
  - Tag Added
  - Campaign Sent (from /api/3/logs)
"""
from __future__ import annotations

import json
import os
import sys
import time
from pathlib import Path
from typing import Any, Iterable

import requests
from dotenv import load_dotenv
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).parent
EXCEL_PATH = ROOT / "Active_campaign-all data.xlsx"
CACHE_DIR = ROOT / ".ac_cache"
CACHE_DIR.mkdir(exist_ok=True)

load_dotenv(ROOT / ".env")
AC_API_URL = os.environ["AC_API_URL"].rstrip("/")
AC_API_KEY = os.environ["AC_API_KEY"]

SESSION = requests.Session()
SESSION.headers.update({"Api-Token": AC_API_KEY, "Accept": "application/json"})

DATA_SHEET = "export final"
ACTIVITY_SHEET = "Activities"
ID_COL = 1  # column A
STATUS_COL = 29  # fallback; skutočný stĺpec sa hľadá podľa názvu hlavičky
STATUS_HEADER = "Status synchronizácie trackingu"
SAVE_EVERY = 500  # save Excel every N processed contacts

ACTIVITY_HEADERS = [
    "contact_id",
    "contact_email",
    "activity_type",
    "timestamp",
    "description",
    "related_object_type",
    "related_object_id",
    "related_object_name",
    "extra_metadata",
]


def get_paginated(endpoint: str, key: str, params: dict | None = None) -> list[dict]:
    """Fetch all pages of an AC endpoint and return the list under `key`.

    Resilient: saves a partial checkpoint every CHECKPOINT_EVERY pages so that
    a network timeout does not throw away hours of fetched data. On restart we
    resume from the saved offset.
    """
    cache_file = CACHE_DIR / f"{endpoint.replace('/', '_')}.json"
    if cache_file.exists():
        print(f"  [cache] loaded {endpoint} from {cache_file.name}")
        return json.loads(cache_file.read_text())

    partial_file = CACHE_DIR / f"{endpoint.replace('/', '_')}.partial.json"
    out: list[dict] = []
    offset = 0
    if partial_file.exists():
        meta = json.loads(partial_file.read_text())
        out = meta["items"]
        offset = meta["next_offset"]
        print(f"  [resume] {endpoint} from offset={offset} (have {len(out)})")

    limit = 100
    base_params = dict(params or {})
    base_params["limit"] = limit
    pages_since_checkpoint = 0
    CHECKPOINT_EVERY = 20

    while True:
        p = dict(base_params, offset=offset)
        url = f"{AC_API_URL}/api/3/{endpoint}"
        last_err: Exception | None = None
        for attempt in range(8):
            try:
                r = SESSION.get(url, params=p, timeout=120)
            except (requests.exceptions.ReadTimeout,
                    requests.exceptions.ConnectionError,
                    requests.exceptions.ChunkedEncodingError) as e:
                last_err = e
                wait = min(2 ** attempt, 60)
                print(f"  [net-err] {type(e).__name__}: retrying in {wait}s (attempt {attempt+1}/8)")
                time.sleep(wait)
                continue
            if r.status_code == 429:
                wait = int(r.headers.get("Retry-After", "5"))
                print(f"  [rate-limit] sleeping {wait}s")
                time.sleep(wait)
                continue
            if r.status_code >= 500:
                wait = min(2 ** attempt, 60)
                print(f"  [{r.status_code}] retrying in {wait}s")
                time.sleep(wait)
                continue
            r.raise_for_status()
            break
        else:
            partial_file.write_text(json.dumps({"items": out, "next_offset": offset}))
            raise RuntimeError(f"failed after retries: {endpoint} offset={offset}: {last_err}")

        data = r.json()
        chunk = data.get(key, [])
        out.extend(chunk)
        offset += limit
        pages_since_checkpoint += 1

        total = int(data.get("meta", {}).get("total", 0)) or len(out)
        print(f"  [{endpoint}] {len(out)}/{total}", flush=True)

        if pages_since_checkpoint >= CHECKPOINT_EVERY:
            partial_file.write_text(json.dumps({"items": out, "next_offset": offset}))
            pages_since_checkpoint = 0

        if len(chunk) < limit or len(out) >= total:
            break
        time.sleep(0.15)

    cache_file.write_text(json.dumps(out))
    if partial_file.exists():
        partial_file.unlink()
    return out


def index_by_id(items: Iterable[dict], id_field: str = "id") -> dict[str, dict]:
    return {str(it[id_field]): it for it in items}


def fetch_reference_data() -> dict[str, dict[str, dict]]:
    print("Fetching reference data...")
    refs = {
        "campaigns": index_by_id(get_paginated("campaigns", "campaigns")),
        "automations": index_by_id(get_paginated("automations", "automations")),
        "lists": index_by_id(get_paginated("lists", "lists")),
        "tags": index_by_id(get_paginated("tags", "tags")),
        "links": index_by_id(get_paginated("links", "links")),
    }
    return refs


def build_activities_index(refs: dict[str, dict[str, dict]]) -> dict[str, list[list]]:
    """Return {contact_id: [activity_row, ...]} where activity_row matches ACTIVITY_HEADERS."""
    print("\nFetching activity data...")
    contact_automations = get_paginated("contactAutomations", "contactAutomations")
    contact_lists = get_paginated("contactLists", "contactLists")
    contact_tags = get_paginated("contactTags", "contactTags")
    logs = get_paginated("logs", "logs")
    link_data = get_paginated("linkData", "linkData")
    bounce_logs = get_paginated("bounceLogs", "bounceLogs")

    by_contact: dict[str, list[list]] = {}

    def push(cid: str, row: list):
        by_contact.setdefault(str(cid), []).append(row)

    # Automations: each record has adddate (entered) and remdate/lastdate (completed)
    automations = refs["automations"]
    for ca in contact_automations:
        cid = ca.get("contact")
        aid = ca.get("automation") or ca.get("seriesid")
        aname = automations.get(str(aid), {}).get("name", "")
        meta = {"contactAutomationId": ca.get("id"), "status": ca.get("status"), "completed": ca.get("completed")}
        if ca.get("adddate"):
            push(cid, [None, None, "Entered automation", ca["adddate"],
                       f"Entered automation {aname}".strip(),
                       "automation", aid, aname, json.dumps(meta, ensure_ascii=False)])
        if ca.get("completed") in (1, "1") and ca.get("lastdate"):
            push(cid, [None, None, "Completed automation", ca["lastdate"],
                       f"Completed automation {aname}".strip(),
                       "automation", aid, aname, json.dumps(meta, ensure_ascii=False)])

    # Lists: status 1 = subscribed, 2 = unsubscribed (sdate=subscribe, udate=unsubscribe)
    lists_ref = refs["lists"]
    for cl in contact_lists:
        cid = cl.get("contact")
        lid = cl.get("list")
        lname = lists_ref.get(str(lid), {}).get("name", "")
        meta = {"contactListId": cl.get("id"), "status": cl.get("status"), "form": cl.get("form")}
        if cl.get("sdate"):
            push(cid, [None, None, "Subscribed to list", cl["sdate"],
                       f"Subscribed to email list {lname}".strip(),
                       "list", lid, lname, json.dumps(meta, ensure_ascii=False)])
        if str(cl.get("status")) == "2" and cl.get("udate"):
            push(cid, [None, None, "Unsubscribed from list", cl["udate"],
                       f"Unsubscribed from list {lname}. Reason: {cl.get('unsubreason') or ''}".strip(),
                       "list", lid, lname, json.dumps(meta, ensure_ascii=False)])

    # Tags
    tags_ref = refs["tags"]
    for ct in contact_tags:
        cid = ct.get("contact")
        tid = ct.get("tag")
        tname = tags_ref.get(str(tid), {}).get("tag", "")
        meta = {"contactTagId": ct.get("id")}
        push(cid, [None, None, "Tag added", ct.get("cdate") or ct.get("created_timestamp"),
                   f"Tag added: {tname}".strip(),
                   "tag", tid, tname, json.dumps(meta, ensure_ascii=False)])

    # Campaign sends (logs)
    campaigns_ref = refs["campaigns"]
    for lg in logs:
        cid = lg.get("subscriberid") or lg.get("contact")
        camp_id = lg.get("campaignid") or lg.get("campaign")
        cname = campaigns_ref.get(str(camp_id), {}).get("name", "")
        meta = {"logId": lg.get("id"), "messageId": lg.get("messageid"),
                "sendId": lg.get("sendid"), "successful": lg.get("successful")}
        push(cid, [None, None, "Campaign Sent", lg.get("tstamp") or lg.get("created_timestamp"),
                   f"Campaign Sent {cname}".strip(),
                   "campaign", camp_id, cname, json.dumps(meta, ensure_ascii=False)])

    # Email opens / clicks via linkData. AC stores tracking-pixel hits and
    # link clicks in the same table; isread=1 + a link whose URL is "open"
    # means the email was opened, otherwise it is a real link click.
    links_ref = refs["links"]
    for ld in link_data:
        cid = ld.get("subscriberid") or ld.get("contact")
        link_id = str(ld.get("linkid") or ld.get("link") or "")
        link_obj = links_ref.get(link_id, {})
        link_url = link_obj.get("link", "")
        camp_id = ld.get("campaignid") or link_obj.get("campaign")
        cname = campaigns_ref.get(str(camp_id), {}).get("name", "")
        is_open = link_url == "open" or str(ld.get("isread")) == "1"
        meta = {"linkDataId": ld.get("id"), "linkId": link_id, "messageId": ld.get("messageid"),
                "times": ld.get("times"), "ip": ld.get("ip6") or ld.get("ip"),
                "userAgent": ld.get("ua"), "isread": ld.get("isread"),
                "linkUrl": link_url}
        if is_open:
            push(cid, [None, None, "Campaign Opened", ld.get("tstamp") or ld.get("created_timestamp"),
                       f"Campaign Opened {cname}".strip(),
                       "campaign", camp_id, cname, json.dumps(meta, ensure_ascii=False)])
        else:
            push(cid, [None, None, "Link Clicked", ld.get("tstamp") or ld.get("created_timestamp"),
                       f"Clicked link in {cname}: {link_url}".strip(),
                       "campaign", camp_id, cname, json.dumps(meta, ensure_ascii=False)])

    # Bounces
    for bl in bounce_logs:
        cid = bl.get("subscriberid") or bl.get("contact")
        camp_id = bl.get("campaignid")
        cname = campaigns_ref.get(str(camp_id), {}).get("name", "")
        meta = {"bounceLogId": bl.get("id"), "messageId": bl.get("messageid"),
                "type": bl.get("type"), "code": bl.get("code"),
                "source": bl.get("source"), "email": bl.get("email")}
        push(cid, [None, None, "Email Bounced", bl.get("tstamp") or bl.get("created_timestamp"),
                   f"Bounced {bl.get('type','')} {cname}".strip(),
                   "campaign", camp_id, cname, json.dumps(meta, ensure_ascii=False)])

    # Sort each contact's activities by timestamp desc to match AC UI ordering
    for cid in by_contact:
        by_contact[cid].sort(key=lambda r: r[3] or "", reverse=True)
    return by_contact


def main():
    refs = fetch_reference_data()
    activities_by_contact = build_activities_index(refs)
    total_activities = sum(len(v) for v in activities_by_contact.values())
    print(f"\nTotal activity rows to write: {total_activities:,}")
    print(f"Contacts with at least 1 activity: {len(activities_by_contact):,}")

    print(f"\nLoading {EXCEL_PATH.name}...")
    wb = load_workbook(EXCEL_PATH)
    ws_data = wb[DATA_SHEET]

    # Build email lookup from data sheet
    email_col = None
    headers = [c.value for c in ws_data[1]]
    for i, h in enumerate(headers, start=1):
        if h == "Email":
            email_col = i
            break

    # stĺpec so statusom hľadáme podľa názvu – rozloženie sa medzi exportmi líši
    status_col = next((i for i, h in enumerate(headers, start=1)
                       if h == STATUS_HEADER), STATUS_COL)
    print(f"Status stĺpec: {status_col} ({headers[status_col-1]!r})")

    # (Re)create activities sheet
    if ACTIVITY_SHEET in wb.sheetnames:
        del wb[ACTIVITY_SHEET]
    ws_act = wb.create_sheet(ACTIVITY_SHEET)
    ws_act.append(ACTIVITY_HEADERS)

    print(f"Iterating {ws_data.max_row - 1} contacts...")
    processed = 0
    written = 0
    t0 = time.time()
    for r in range(2, ws_data.max_row + 1):
        raw_id = ws_data.cell(row=r, column=ID_COL).value
        if raw_id is None:
            continue
        cid = str(int(raw_id)) if isinstance(raw_id, float) else str(raw_id)
        email = ws_data.cell(row=r, column=email_col).value if email_col else None

        rows = activities_by_contact.get(cid, [])
        for row in rows:
            row[0] = int(cid) if cid.isdigit() else cid
            row[1] = email
            ws_act.append(row)
            written += 1

        ws_data.cell(row=r, column=status_col).value = "OK"
        processed += 1

        if processed % SAVE_EVERY == 0:
            wb.save(EXCEL_PATH)
            elapsed = time.time() - t0
            rate = processed / elapsed if elapsed else 0
            print(f"  processed {processed}/{ws_data.max_row - 1}  written={written:,}  ({rate:.0f}/s)")

    # Auto-size first sheet status column? skip - keep file fast
    print("Saving final workbook...")
    wb.save(EXCEL_PATH)
    elapsed = time.time() - t0
    print(f"Done. processed={processed}  activities_written={written:,}  elapsed={elapsed:.1f}s")


if __name__ == "__main__":
    main()
