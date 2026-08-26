"""
Stiahne kontakty z ActiveCampaign cez API a zapíše hárok 'export final'
v rovnakom stĺpcovom rozložení, aké mal pôvodný ručný export z AC UI.

LEN ČÍTANIE. Používa výhradne GET /api/3/... — nič nemaže ani nemení.

Vytvorí Active_campaign-all data.xlsx s hárkom 'export final'; potom sa naň
dá pustiť export_activities.py, ktorý doplní hárok 'Activities'.
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

from openpyxl import Workbook

# helpery (paginácia, retry, cache) preberáme z existujúceho skriptu
from export_activities import get_paginated, CACHE_DIR, ROOT

OUT_PATH = ROOT / "Active_campaign-all data.xlsx"
SHEET = "export final"

# presné poradie stĺpcov z pôvodného exportu
BUILTIN = [
    ("ID", "id"),
    ("Email", "email"),
    ("First Name", "firstName"),
    ("Last Name", "lastName"),
    ("Phone Number", "phone"),
    ("Date Created", "cdate"),
    ("IP Address", "ip"),
    ("User Agent", "ua"),
]
# custom polia v pôvodnom poradí; hodnota = názov (title) poľa v AC
CUSTOM = [
    "PSC", "utm_campaign", "utm_medium", "utm_source", "utm_term",
    "Hlavná otázka č.1", "Podotázka č.1", "Hlavná otázka č.2", "Podotázka č.2",
    "Hlavná otázka č.3", "Podotázka č.3", "Hlavná otázka č.4",
    "bnrcode", "Mesto",
]
# lead scores (AC "Scores") – v pôvodnom exporte to boli stĺpce s hviezdičkou
SCORES = ["Email Activity", "Website Scoring", "Score 3"]
TRAILING = ["Tags", "Status synchronizácie trackingu"]



def fmt_ip(v):
    """AC drží IP ako 32-bit integer; pôvodný UI export ju mal v bodkovom tvare."""
    if v in (None, "", "0"):
        return None
    try:
        n = int(v)
    except (TypeError, ValueError):
        return v
    if n <= 0:
        return None
    return ".".join(str((n >> sh) & 255) for sh in (24, 16, 8, 0))


def fmt_dt(v):
    """'2026-07-24T09:03:11-05:00' -> '2026-07-24 9:03:11' (ako pôvodný export)."""
    if not v:
        return None
    t = str(v).replace("T", " ")
    for cut in ("+", "Z"):
        if cut in t[10:]:
            t = t[:10] + t[10:].split(cut)[0]
    if t.count("-") > 2:
        head, _, tail = t.rpartition("-")
        if ":" in tail:
            t = head
    t = t.strip()
    d, _, tm = t.partition(" ")
    if tm and tm[0] == "0":
        tm = tm[1:]
    return f"{d} {tm}".strip()


def norm(s: str) -> str:
    return "".join(str(s or "").lower().split()).replace("*", "")


def main():
    print("Sťahujem kontakty a polia (len GET)...")
    fields = get_paginated("fields", "fields")
    contacts = get_paginated("contacts", "contacts")
    field_values = get_paginated("fieldValues", "fieldValues")
    contact_tags = get_paginated("contactTags", "contactTags")
    tags = get_paginated("tags", "tags")
    scores = get_paginated("scores", "scores")
    score_values = get_paginated("scoreValues", "scoreValues")

    print(f"\nkontaktov: {len(contacts):,} | polí: {len(fields)} | "
          f"hodnôt polí: {len(field_values):,} | tagov na kontaktoch: {len(contact_tags):,}")

    # mapovanie názov custom poľa -> id
    by_norm = {norm(f.get("title")): str(f["id"]) for f in fields}
    print("\nCustom polia v účte:")
    for f in sorted(fields, key=lambda x: int(x["id"])):
        print(f"   id={f['id']:>4}  {f.get('title')!r}  ({f.get('type')})")

    col_field_id = {}
    missing = []
    for title in CUSTOM:
        fid = by_norm.get(norm(title))
        if fid:
            col_field_id[title] = fid
        else:
            missing.append(title)
    if missing:
        print(f"\n!! POZOR: tieto polia z pôvodného exportu v účte nie sú: {missing}")
        print("   (stĺpce zostanú prázdne — účet má zrejme inú sadu polí)")

    # lead scores: názov -> id, a potom {contact_id: {score_id: hodnota}}
    score_id = {norm(sc.get("name")): str(sc["id"]) for sc in scores}
    print("\nLead scores v účte:", [sc.get("name") for sc in scores])
    svals: dict[str, dict[str, str]] = {}
    for sv in score_values:
        svals.setdefault(str(sv.get("contact")), {})[str(sv.get("score"))] = sv.get("scoreValue")

    # hodnoty polí: {contact_id: {field_id: value}}
    vals: dict[str, dict[str, str]] = {}
    for fv in field_values:
        vals.setdefault(str(fv.get("contact")), {})[str(fv.get("field"))] = fv.get("value")

    tag_name = {str(t["id"]): t.get("tag", "") for t in tags}
    ctags: dict[str, list[str]] = {}
    for ct in contact_tags:
        ctags.setdefault(str(ct.get("contact")), []).append(tag_name.get(str(ct.get("tag")), ""))

    wb = Workbook()
    ws = wb.active
    ws.title = SHEET
    header = ([h for h, _ in BUILTIN]
              + [f"*{t}" for t in CUSTOM]
              + [f"*{t}" for t in SCORES]
              + TRAILING)
    ws.append(header)

    for c in contacts:
        cid = str(c.get("id"))
        row = []
        for label, key in BUILTIN:
            v = c.get(key)
            if key == "ip":
                v = fmt_ip(v)
            elif key == "cdate":
                v = fmt_dt(v)
            row.append(v)
        for title in CUSTOM:
            fid = col_field_id.get(title)
            row.append(vals.get(cid, {}).get(fid) if fid else None)
        for title in SCORES:
            sid = score_id.get(norm(title))
            row.append(svals.get(cid, {}).get(sid) if sid else None)
        row.append(", ".join(sorted(t for t in ctags.get(cid, []) if t)))
        row.append(None)  # Status synchronizácie trackingu – doplní export_activities.py
        ws.append(row)

    wb.save(OUT_PATH)
    print(f"\nHotovo: {OUT_PATH.name}  ({ws.max_row - 1:,} kontaktov x {len(header)} stĺpcov)")


if __name__ == "__main__":
    main()
