"""
D365 Drill Engine — Dataverse Push
====================================
Reads enriched_cases.jsonl and pushes new records to Dataverse
mwi_drillcard table via OData API.

Usage:
  python dataverse_push.py              # push all new cases
  python dataverse_push.py --dry-run    # show what would be pushed
  python dataverse_push.py --replace    # delete all and re-push everything

Run from scrapper/ folder:
  cd scrapper
  python dataverse_push.py
"""

import os
import json
import time
import argparse
import logging
from pathlib import Path
from dotenv import load_dotenv
import msal
import requests

load_dotenv(dotenv_path=os.path.join(os.path.dirname(__file__), '..', '.env'))

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-7s  %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("dataverse_push")

# ── Config ──────────────────────────────────────────────────────
DATAVERSE_URL  = os.getenv("DATAVERSE_URL")
TENANT_ID      = os.getenv("DRILL_TENANT_ID")
CLIENT_ID      = os.getenv("DRILL_CLIENT_ID")
CLIENT_SECRET  = os.getenv("DRILL_CLIENT_SECRET")

ENTITY         = "cdr_drillcards"
API_URL        = f"{DATAVERSE_URL}/api/data/v9.2/{ENTITY}"
SCOPE          = f"{DATAVERSE_URL}/.default"

INPUT_FILE = Path(os.path.join(os.path.dirname(__file__), '..', 'scrapper', 'output', 'enriched_cases.jsonl'))
RATE_LIMIT     = 0.5   # seconds between requests


# ── Auth ─────────────────────────────────────────────────────────

def get_token() -> str:
    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        client_credential=CLIENT_SECRET,
    )
    result = app.acquire_token_for_client(scopes=[SCOPE])
    if "access_token" not in result:
        raise RuntimeError(f"Token error: {result.get('error_description')}")
    return result["access_token"]


def headers(token: str) -> dict:
    return {
        "Authorization": f"Bearer {token}",
        "Content-Type":  "application/json",
        "OData-MaxVersion": "4.0",
        "OData-Version":    "4.0",
        "Prefer": "return=representation",
    }


# ── Existing records ─────────────────────────────────────────────

def get_existing_case_ids(token: str) -> set:
    """Fetch all existing CaseID values from Dataverse."""
    url = f"{API_URL}?$select=cdr_caseid&$top=1000"
    resp = requests.get(url, headers=headers(token), timeout=15)
    resp.raise_for_status()
    records = resp.json().get("value", [])
    return {r.get("cdr_caseid", "") for r in records if r.get("cdr_caseid")}


def delete_all(token: str):
    """Delete all existing drill card records."""
    url = f"{API_URL}?$select=cdr_drillcardid&$top=1000"
    resp = requests.get(url, headers=headers(token), timeout=15)
    resp.raise_for_status()
    records = resp.json().get("value", [])

    log.info(f"Deleting {len(records)} existing records...")
    for r in records:
        rid = r["cdr_drillcardid"]
        del_resp = requests.delete(
            f"{API_URL}({rid})",
            headers=headers(token),
            timeout=15
        )
        if del_resp.status_code == 204:
            log.info(f"  Deleted: {rid[:8]}")
        time.sleep(RATE_LIMIT)


# ── Field mapping ─────────────────────────────────────────────────

def to_dataverse_record(case: dict) -> dict:
    """Map enriched case dict to Dataverse column names."""
    # Answer Is MSFT — convert to boolean
    answer_is_msft = case.get("answer_is_msft", False)
    if isinstance(answer_is_msft, str):
        answer_is_msft = answer_is_msft.lower() in ("yes", "true", "1")

    # Quality score — ensure float
    try:
        quality_score = float(case.get("quality_score", 0))
    except (ValueError, TypeError):
        quality_score = 0.0

    # Case date — prefer thread_date, fall back to scraped_at
    thread_date = case.get("thread_date", "")
    scraped_at = case.get("scraped_at", "")
    case_date = thread_date[:10] if thread_date else (scraped_at[:10] if scraped_at else None)

    # Validate date — Dataverse minimum is 1753-01-01
    if case_date:
        try:
            year = int(case_date[:4])
            if year < 1753 or year > 9999:
                case_date = None
        except (ValueError, TypeError):
            case_date = None

    return {
        "cdr_symptom":         str(case.get("symptom", ""))[:300],
        "cdr_caseid":          str(case.get("case_id", ""))[:100],
        "cdr_casedate":        case_date,
        "cdr_module":          str(case.get("module", ""))[:100],
        "cdr_category":        str(case.get("category", ""))[:100],
        "cdr_rootcause":       str(case.get("root_cause", ""))[:2000],
        "cdr_diagnosticsteps": str(case.get("diagnostic_steps", ""))[:4000],
        "cdr_d365navigation":  str(case.get("d365_navigation", ""))[:2000],
        "cdr_resolution":      str(case.get("resolution", ""))[:2000],
        "cdr_difficulty":      str(case.get("difficulty", ""))[:50],
        "cdr_consultanttip":   str(case.get("consultant_tip", ""))[:2000],
        "cdr_relatedconcepts": str(case.get("related_concepts", ""))[:500],
        "cdr_qualityscore":    quality_score,
        "cdr_answerismsft":    answer_is_msft,
        "cdr_sourceurl":       str(case.get("source_url", ""))[:500],
        "cdr_learnurl":        str(case.get("learn_url", ""))[:500],
        "cdr_rawtitle":        str(case.get("raw_title", ""))[:300],
        "cdr_scrapedat":       str(case.get("scraped_at", ""))[:100],
    }


# ── Push ──────────────────────────────────────────────────────────

def push_case(record: dict, token: str) -> bool:
    """Create a single drill card record in Dataverse."""
    resp = requests.post(
        API_URL,
        headers=headers(token),
        json=record,
        timeout=15,
    )
    if resp.status_code in (200, 201, 204):
        return True
    else:
        log.warning(f"  Push failed {resp.status_code}: {resp.text[:200]}")
        return False


# ── Main ──────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--dry-run",  action="store_true")
    parser.add_argument("--replace",  action="store_true",
                        help="Delete all existing records and re-push everything")
    args = parser.parse_args()

    # Validate config
    if not DATAVERSE_URL or not CLIENT_ID:
        log.error("Missing env vars. Check DATAVERSE_URL, DRILL_CLIENT_ID in .env")
        return

    # Load local cases
    if not INPUT_FILE.exists():
        log.error(f"Input file not found: {INPUT_FILE}")
        log.error("Run scraper.py first to generate enriched_cases.jsonl")
        return

    cases = []
    with open(INPUT_FILE, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line:
                try:
                    cases.append(json.loads(line))
                except json.JSONDecodeError:
                    pass

    log.info(f"Loaded {len(cases)} cases from {INPUT_FILE}")

    if args.dry_run:
        log.info("Dry run — showing what would be pushed:")
        for c in cases:
            log.info(f"  [{c.get('case_id','')}] {c.get('symptom','')[:60]}")
        return

    # Get token
    log.info("Authenticating with Dataverse...")
    token = get_token()
    log.info("✓ Token acquired")

    # Handle replace mode
    if args.replace:
        delete_all(token)
        existing_ids = set()
    else:
        log.info("Checking existing records...")
        existing_ids = get_existing_case_ids(token)
        log.info(f"  Found {len(existing_ids)} existing records")

    # Push new cases
    pushed = 0
    skipped = 0
    errors = 0

    for i, case in enumerate(cases, 1):
        case_id = case.get("case_id", "")

        if case_id in existing_ids:
            log.info(f"[{i}/{len(cases)}] Skip (exists): {case_id}")
            skipped += 1
            continue

        record = to_dataverse_record(case)
        symptom = case.get("symptom", "")[:55]

        log.info(f"[{i}/{len(cases)}] Pushing: [{case_id}] {symptom}...")

        if push_case(record, token):
            log.info(f"  ✓ Pushed")
            pushed += 1
        else:
            errors += 1

        time.sleep(RATE_LIMIT)

    # Summary
    print(f"\n{'='*50}")
    print(f"PUSH COMPLETE")
    print(f"  Pushed:   {pushed}")
    print(f"  Skipped:  {skipped}  (already in Dataverse)")
    print(f"  Errors:   {errors}")
    print(f"{'='*50}")


if __name__ == "__main__":
    main()
