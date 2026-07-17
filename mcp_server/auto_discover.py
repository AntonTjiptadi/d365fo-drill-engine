"""
D365 Drill Engine — Auto Discovery
====================================
Uses SerpAPI to search Google for D365 Production Control and Master Planning
forum threads on community.dynamics.com.

Finds new thread IDs not already in enriched_cases.jsonl,
runs the scraper on them, then pushes to Dataverse.

Usage:
  python auto_discover.py                    # discover + enrich + push, limit 20
  python auto_discover.py --dry-run          # show what would be found, no processing
  python auto_discover.py --limit 10         # cap new threads per run
  python auto_discover.py --category "Production Control"  # focus on one category

Run from mcp_server/ folder:
  cd mcp_server
  python auto_discover.py
"""

import os
import re
import sys
import json
import time
import logging
import argparse
import subprocess
from pathlib import Path
from datetime import datetime
from dotenv import load_dotenv
import requests

load_dotenv(dotenv_path=os.path.join(os.path.dirname(__file__), '..', '.env'))

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-7s  %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("auto_discover")

# ── Config ──────────────────────────────────────────────────────
SERP_API_KEY   = os.getenv("SERP_API_KEY")
SCRAPPER_DIR   = Path(os.path.dirname(__file__)) / ".." / "scrapper"
ENRICHED_FILE  = SCRAPPER_DIR / "output" / "enriched_cases.jsonl"
SKIPPED_FILE   = SCRAPPER_DIR / "output" / "skipped.jsonl"

RATE_LIMIT     = 1.5   # seconds between SerpAPI calls
STATUS_FILE    = Path(os.path.dirname(__file__)) / "discover_status.json"


# ── Status file helpers ──────────────────────────────────────────

def write_status(stage: str, detail: str = "", **kwargs):
    """Write current pipeline progress to discover_status.json."""
    status = {
        "stage":      stage,
        "detail":     detail,
        "updated_at": datetime.utcnow().isoformat(),
        **kwargs,
    }
    try:
        with open(STATUS_FILE, "w", encoding="utf-8") as f:
            json.dump(status, f, indent=2)
    except Exception:
        pass

# ── Search queries ───────────────────────────────────────────────
# Each query targets a specific symptom area
# site: restricts to community.dynamics.com only
SEARCH_QUERIES = [
    # Master Planning
    "site:community.dynamics.com planned order not generated D365 F&O",
    "site:community.dynamics.com coverage group master planning issue D365",
    "site:community.dynamics.com MRP no planned orders supply chain",
    "site:community.dynamics.com negative days master planning D365",
    "site:community.dynamics.com planning optimization planned orders",
    "site:community.dynamics.com net requirements empty D365 F&O",
    "site:community.dynamics.com demand forecast master planning D365",
    "site:community.dynamics.com item coverage setup D365 supply chain",
    "site:community.dynamics.com firming planned orders D365",
    "site:community.dynamics.com safety stock master planning D365",

    # Production Control
    "site:community.dynamics.com production order cannot estimate D365",
    "site:community.dynamics.com route version not approved D365",
    "site:community.dynamics.com report as finished error D365 F&O",
    "site:community.dynamics.com picking list not registered production",
    "site:community.dynamics.com BOM consumption wrong D365 production",
    "site:community.dynamics.com flushing principle not working D365",
    "site:community.dynamics.com production order status stuck D365",
    "site:community.dynamics.com capacity overload scheduling D365",
    "site:community.dynamics.com production variance D365 F&O",
    "site:community.dynamics.com job scheduling D365 production",

    # Cross-cutting
    "site:community.dynamics.com lead time mismatch production D365",
    "site:community.dynamics.com reservation failure production order D365",
    "site:community.dynamics.com production costing standard cost D365",
    "site:community.dynamics.com resource calendar capacity D365",
    "site:community.dynamics.com subcontracting production D365 F&O",
]


# ── Load already processed IDs ───────────────────────────────────
def load_existing_ids() -> set:
    """Load already-processed thread IDs from both JSONL and Dataverse."""
    ids = set()

    # From local JSONL files
    for path in [ENRICHED_FILE, SKIPPED_FILE]:
        if path.exists():
            with open(path, encoding="utf-8") as f:
                for line in f:
                    try:
                        rec = json.loads(line.strip())
                        case_id = rec.get("case_id", "")
                        url = rec.get("source_url", "")
                        if case_id:
                            ids.add(case_id.replace("PC-", "").lower()[:8])
                        match = re.search(r"threadid=([a-f0-9\-]{8,})", url, re.I)
                        if match:
                            ids.add(match.group(1)[:8].lower())
                    except Exception:
                        pass

    log.info(f"  From local JSONL: {len(ids)} IDs")

    # From Dataverse — fetch all existing case IDs
    try:
        import msal
        tenant_id     = os.getenv("DRILL_TENANT_ID")
        client_id     = os.getenv("DRILL_CLIENT_ID")
        client_secret = os.getenv("DRILL_CLIENT_SECRET")
        dataverse_url = os.getenv("DATAVERSE_URL")

        app = msal.ConfidentialClientApplication(
            client_id,
            authority=f"https://login.microsoftonline.com/{tenant_id}",
            client_credential=client_secret,
        )
        token = app.acquire_token_for_client(
            scopes=[f"{dataverse_url}/.default"]
        )
        headers = {
            "Authorization": f"Bearer {token['access_token']}",
            "OData-MaxVersion": "4.0",
            "OData-Version": "4.0",
            "Accept": "application/json",
        }
        r = requests.get(
            f"{dataverse_url}/api/data/v9.2/cdr_drillcards"
            f"?$select=cdr_caseid&$top=1000",
            headers=headers,
            timeout=15,
        )
        r.raise_for_status()
        for rec in r.json().get("value", []):
            case_id = rec.get("cdr_caseid", "")
            if case_id:
                ids.add(case_id.replace("PC-", "").lower()[:8])
        log.info(f"  From Dataverse:   {len(ids)} IDs total")
    except Exception as e:
        log.warning(f"  Could not load from Dataverse: {e}")

    return ids


# ── SerpAPI search ───────────────────────────────────────────────

def serp_search(query: str, num: int = 10) -> list[dict]:
    """
    Search Google via SerpAPI.
    Returns list of organic results with title, link, snippet.
    """
    if not SERP_API_KEY:
        raise RuntimeError("SERP_API_KEY not set in .env")

    params = {
        "q":       query,
        "api_key": SERP_API_KEY,
        "num":     num,
        "engine":  "google",
    }

    resp = requests.get(
        "https://serpapi.com/search",
        params=params,
        timeout=15,
    )
    resp.raise_for_status()
    data = resp.json()

    return data.get("organic_results", [])


def extract_thread_ids(results: list[dict]) -> list[dict]:
    """
    Extract thread IDs from SerpAPI results.
    Returns list of {thread_id, title, url, snippet}
    """
    threads = []
    seen = set()

    for r in results:
        link = r.get("link", "")
        title = r.get("title", "")
        snippet = r.get("snippet", "")
        # SerpAPI often includes a date field — normalise to YYYY-MM-DD
        raw_date = r.get("date", "")
        thread_date = ""
        if raw_date:
            dm = re.search(r"(\d{4}-\d{2}-\d{2})", raw_date)
            if dm:
                thread_date = dm.group(1)

        # Must be a community.dynamics.com thread URL
        if "community.dynamics.com" not in link:
            continue

        match = re.search(r"threadid=([a-f0-9\-]{30,})", link, re.I)
        if not match:
            continue

        thread_id = match.group(1)
        if thread_id in seen:
            continue

        # Skip obviously unrelated content
        title_lower = title.lower()
        skip_keywords = [
            "business central", "nav ", "crm", "customer service",
            "power apps", "power bi", "field service", "marketing",
            "manoverse", "commerce",
        ]
        if any(kw in title_lower for kw in skip_keywords):
            continue

        threads.append({
            "thread_id":    thread_id,
            "title":        title,
            "url":          link,
            "snippet":      snippet[:200],
            "thread_date":  thread_date,
        })
        seen.add(thread_id)

    return threads


# ── Main pipeline ─────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Auto-discover D365 forum threads")
    parser.add_argument("--dry-run",  action="store_true",
                        help="Show what would be found, no processing")
    parser.add_argument("--limit",    type=int, default=20,
                        help="Max new threads to process per run (default: 20)")
    parser.add_argument("--category", type=str, default=None,
                        help="Focus on specific category: 'Master Planning' or 'Production Control'")
    args = parser.parse_args()

    if not SERP_API_KEY:
        log.error("SERP_API_KEY not set in .env")
        return

    log.info("Loading existing thread IDs...")
    existing_ids = load_existing_ids()
    log.info(f"  Already processed: {len(existing_ids)} thread ID prefixes")
    write_status("searching", "Loading existing IDs complete", existing=len(existing_ids))

    # Filter queries by category if specified
    queries = SEARCH_QUERIES
    if args.category:
        cat_lower = args.category.lower()
        if "master" in cat_lower:
            queries = SEARCH_QUERIES[:10]
            log.info("  Focusing on Master Planning queries")
        elif "production" in cat_lower:
            queries = SEARCH_QUERIES[10:20]
            log.info("  Focusing on Production Control queries")

    # ── Discover new threads ──
    log.info(f"\nSearching {len(queries)} queries via SerpAPI...")
    all_new_threads = []
    seen_ids = set()

    for i, query in enumerate(queries, 1):
        log.info(f"  [{i}/{len(queries)}] {query[:70]}")

        try:
            results = serp_search(query, num=10)
            threads = extract_thread_ids(results)

            for t in threads:
                tid_prefix = t["thread_id"][:8].lower()
                if tid_prefix not in existing_ids and tid_prefix not in seen_ids:
                    all_new_threads.append(t)
                    seen_ids.add(tid_prefix)
                    log.info(f"    + NEW: {t['title'][:60]}")

        except Exception as e:
            log.warning(f"  Search failed for query {i}: {e}")

        time.sleep(RATE_LIMIT)
        write_status("searching", f"Query {i}/{len(queries)} done", found_so_far=len(all_new_threads))

        # Stop if we have enough
        if len(all_new_threads) >= args.limit:
            log.info(f"  Reached limit of {args.limit} new threads — stopping search")
            break

    log.info(f"\nFound {len(all_new_threads)} new threads")
    write_status("discovered", f"Found {len(all_new_threads)} new threads", found=len(all_new_threads))

    if not all_new_threads:
        log.info("No new threads found. Knowledge base is up to date.")
        return

    if args.dry_run:
        log.info("\nDry run — threads that would be processed:")
        for t in all_new_threads:
            log.info(f"  [{t['thread_id'][:8]}] {t['title'][:70]}")
        return

    # ── Cap to limit ──
    threads_to_process = all_new_threads[:args.limit]
    log.info(f"\nProcessing {len(threads_to_process)} new threads...")

    # ── Add to scraper seed list and run scraper ──
    # Write thread IDs to a temp file for the scraper to pick up
    temp_seed_file = SCRAPPER_DIR / "temp_new_threads.json"
    with open(temp_seed_file, "w", encoding="utf-8") as f:
        json.dump(threads_to_process, f, indent=2)

    log.info("Running scraper on new threads...")
    write_status("scraping", f"Running scraper on {len(threads_to_process)} threads", to_process=len(threads_to_process))

    scraper_script = SCRAPPER_DIR / "scraper.py"
    python_exe = sys.executable

    result = subprocess.run(
        [python_exe, str(scraper_script), "--seed-file", str(temp_seed_file)],
        capture_output=True,
        text=True,
        cwd=str(SCRAPPER_DIR),
    )

    if result.stdout:
        print(result.stdout)
    if result.stderr:
        log.warning(result.stderr[-500:])

    # Cleanup temp file
    if temp_seed_file.exists():
        temp_seed_file.unlink()

    # ── Push to Dataverse ──
    log.info("\nPushing new cards to Dataverse...")
    write_status("pushing", "Pushing enriched cards to Dataverse")

    push_script = Path(os.path.dirname(__file__)) / "dataverse_push.py"
    result = subprocess.run(
        [python_exe, str(push_script)],
        capture_output=True,
        text=True,
        cwd=str(Path(os.path.dirname(__file__))),
    )

    if result.stdout:
        print(result.stdout)
    if result.stderr:
        log.warning(result.stderr[-300:])

    log.info("\nAuto-discovery complete.")

    # Parse final stats from push output
    import re as _re
    push_out = result.stdout + result.stderr
    def _int(pat):
        m = _re.search(pat, push_out)
        return int(m.group(1)) if m else 0

    write_status(
        "complete",
        "Auto-discovery pipeline finished",
        pushed=_int(r"Pushed:\s+(\d+)"),
        skipped=_int(r"Skipped:\s+(\d+)"),
        errors=_int(r"Errors:\s+(\d+)"),
    )


if __name__ == "__main__":
    main()
