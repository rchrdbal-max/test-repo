"""
Iteratively delete every contact in the AC account.

Strategy: list the first 100 contacts (id ASC), delete each, repeat. Once the
list comes back empty, we are done. Permanent errors (non-404/429 4xx) are
recorded in delete_progress.json so we don't loop on them.

DESTRUCTIVE — DELETE /api/3/contacts/{id} is irreversible.
"""
from __future__ import annotations

import json
import os
import sys
import time
from datetime import datetime
from pathlib import Path

import requests
from dotenv import load_dotenv

ROOT = Path(__file__).parent
PROGRESS_FILE = ROOT / "delete_progress.json"

load_dotenv(ROOT / ".env")
AC_API_URL = os.environ["AC_API_URL"].rstrip("/")
AC_API_KEY = os.environ["AC_API_KEY"]

SESSION = requests.Session()
SESSION.headers.update({"Api-Token": AC_API_KEY, "Accept": "application/json"})

SLEEP_BETWEEN = 0.2  # 5 req/s pacing


def load_progress() -> dict:
    if PROGRESS_FILE.exists():
        return json.loads(PROGRESS_FILE.read_text())
    return {"deleted": 0, "errors": {}}


def save_progress(p: dict) -> None:
    PROGRESS_FILE.write_text(json.dumps(p, indent=2))


def list_first_page(skip_ids: set[str]) -> list[dict]:
    """Return first page of contacts (id ASC) excluding IDs in skip_ids."""
    for attempt in range(6):
        try:
            r = SESSION.get(
                f"{AC_API_URL}/api/3/contacts",
                params={"limit": 100, "orders[id]": "ASC"},
                timeout=120,
            )
        except (requests.exceptions.ReadTimeout,
                requests.exceptions.ConnectionError,
                requests.exceptions.ChunkedEncodingError) as e:
            print(f"  [list net-err] {type(e).__name__}: retry in {min(2**attempt, 30)}s", flush=True)
            time.sleep(min(2 ** attempt, 30))
            continue
        if r.status_code == 429:
            wait = int(r.headers.get("Retry-After", "5"))
            print(f"  [list 429] wait {wait}s", flush=True)
            time.sleep(wait)
            continue
        if r.status_code >= 500:
            print(f"  [list {r.status_code}] retry in {min(2**attempt, 30)}s", flush=True)
            time.sleep(min(2 ** attempt, 30))
            continue
        r.raise_for_status()
        contacts = r.json().get("contacts", [])
        return [c for c in contacts if str(c["id"]) not in skip_ids]
    raise RuntimeError("could not list contacts after retries")


def delete_one(cid: str) -> tuple[str, str]:
    """Returns (status, message). status in: deleted, gone, perm_err, transient_err."""
    url = f"{AC_API_URL}/api/3/contacts/{cid}"
    last_err = ""
    for attempt in range(6):
        try:
            r = SESSION.delete(url, timeout=60)
        except (requests.exceptions.ReadTimeout,
                requests.exceptions.ConnectionError,
                requests.exceptions.ChunkedEncodingError) as e:
            last_err = f"{type(e).__name__}"
            time.sleep(min(2 ** attempt, 30))
            continue
        if r.status_code == 200:
            return "deleted", ""
        if r.status_code == 404:
            return "gone", ""
        if r.status_code == 429:
            wait = int(r.headers.get("Retry-After", "5"))
            time.sleep(wait)
            continue
        if r.status_code >= 500:
            last_err = f"HTTP {r.status_code}"
            time.sleep(min(2 ** attempt, 30))
            continue
        return "perm_err", f"HTTP {r.status_code}: {r.text[:120]}"
    return "transient_err", f"after retries: {last_err}"


def main():
    progress = load_progress()
    skip = set(progress["errors"].keys())
    print(f"Starting. Already deleted (per progress): {progress['deleted']}, "
          f"perm-errors to skip: {len(skip)}")

    t0 = time.time()
    iter_no = 0
    while True:
        iter_no += 1
        page = list_first_page(skip)
        if not page:
            print("List empty — done.")
            break

        print(f"[iter {iter_no}] got {len(page)} contacts (first id={page[0]['id']}, last id={page[-1]['id']})", flush=True)
        for c in page:
            cid = str(c["id"])
            status, msg = delete_one(cid)
            if status in ("deleted", "gone"):
                progress["deleted"] += 1
            else:
                progress["errors"][cid] = f"{status}: {msg}"
                skip.add(cid)
                print(f"  [skip {cid}] {status}: {msg}", flush=True)
            time.sleep(SLEEP_BETWEEN)

        save_progress(progress)
        elapsed = time.time() - t0
        rate = progress["deleted"] / elapsed if elapsed else 0
        print(f"  total deleted={progress['deleted']}  errors={len(progress['errors'])}  "
              f"rate={rate:.1f}/s  elapsed={elapsed/60:.1f}min", flush=True)

    elapsed = time.time() - t0
    print(f"Done. deleted={progress['deleted']}  errors={len(progress['errors'])}  elapsed={elapsed/60:.1f}min")


if __name__ == "__main__":
    main()
