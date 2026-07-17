"""
D365 Drill Engine — MCP Server
==============================
FastMCP server exposing 7 tools for the Copilot Studio drill agent.

Tools:
  1. get_random_drill        — fetch one drill card (filtered by category/difficulty)
  2. get_case_by_symptom     — semantic search by symptom keyword
  3. get_cases_by_category   — list all cases in a category
  4. score_answer            — evaluate user's diagnosis against correct answer
  5. get_weak_areas          — analyse session history, return weak categories
  6. get_diagnostic_path     — return full step-by-step diagnostic for a case
  7. get_odata_query         — return the D365 OData verification query for a case
  8. get_learn_reference     — fetch Microsoft Learn content for a case for cross-validation
  9. run_auto_discover       — discover and ingest new forum threads into the knowledge base
  
Transport: HTTP (StreamableHTTP) for Copilot Studio / ngrok
Port: 8002 (avoids conflict with WMS server on 8000/8001)

Usage:
  python server.py
  Then expose via: ngrok http 8002
  Register in Copilot Studio: https://<ngrok-url>/mcp
"""

import os
import random
import json
import logging
from typing import Optional
from dotenv import load_dotenv

load_dotenv(dotenv_path=os.path.join(os.path.dirname(__file__), '..', '.env'))

import anthropic
from fastmcp import FastMCP
# from sharepoint_client import SharePointClient
from dataverse_client import DataverseClient

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger(__name__)

mcp = FastMCP(
    name="D365 Drill Engine",
    instructions=(
        "You are a D365 Production Control and Master Planning drill instructor. "
        "Default drill loop: "
        "1. Call get_random_drill to fetch a card. "
        "2. Present the symptom to the user and ask for their diagnosis. "
        "3. After they answer (or say 'i don't know'), call score_answer. "
        "4. Always call get_learn_reference after scoring to cross-validate. "
        "5. After every 5 drills, call get_weak_areas with session results. "
        "Never reveal the root cause or resolution before the user attempts a diagnosis."
    ),
)

# sp = SharePointClient()
sp = DataverseClient()
# anthropic_client removed — scoring handled by Claude Desktop Pro model


# ─────────────────────────────────────────────
# TOOL 1 — get_random_drill
# ─────────────────────────────────────────────

@mcp.tool()
def get_random_drill(
    module: Optional[str] = None,
    category: Optional[str] = None,
    difficulty: Optional[str] = None,
) -> dict:
    """
    Return one random drill card for the user to diagnose.
    Only returns: case_id, symptom, module, category, difficulty.
    Root cause and resolution are withheld until score_answer is called.

    Args:
        module:     Filter by module e.g. 'Master Planning', 'Production Control'
        category:   Filter by category e.g. 'Coverage Settings', 'Route & Operations'
        difficulty: Filter by difficulty e.g. 'beginner', 'intermediate', 'advanced'
    """
    try:
        cases = sp.get_cases_by_filter(
            module=module,
            category=category,
            difficulty=difficulty,
        )
        if not cases:
            return {"error": "No drill cards found for the given filters."}

        case = random.choice(cases)

        # Return symptom only — never reveal answer before attempt
        return {
            "case_id":    case["case_id"],
            "symptom":    case["symptom"],
            "module":     case["module"],
            "category":   case["category"],
            "difficulty": case["difficulty"],
            "instruction": (
                "Present this symptom to the user and ask them to diagnose the root cause. "
                "Do NOT reveal the answer yet. Wait for their response, then call score_answer."
            ),
        }
    except Exception as e:
        log.error(f"get_random_drill error: {e}")
        return {"error": str(e)}


# ─────────────────────────────────────────────
# TOOL 2 — get_case_by_symptom
# ─────────────────────────────────────────────

@mcp.tool()
def get_case_by_symptom(symptom_text: str) -> list[dict]:
    """
    Search for drill cards matching a symptom keyword.
    Returns up to 5 closest matches with full card content.
    Use this in Lookup Mode when the user asks about a specific issue.

    Args:
        symptom_text: Keyword or phrase describing the symptom e.g. 'planned order not generated'
    """
    try:
        cases = sp.search_by_symptom(symptom_text)
        if not cases:
            return [{"error": f"No cases found matching: {symptom_text}"}]
        return cases
    except Exception as e:
        log.error(f"get_case_by_symptom error: {e}")
        return [{"error": str(e)}]


# ─────────────────────────────────────────────
# TOOL 3 — get_cases_by_category
# ─────────────────────────────────────────────

@mcp.tool()
def get_cases_by_category(category: str) -> list[dict]:
    """
    Return all drill cards in a specific category.
    Use this to give the user an overview of a topic area.

    Args:
        category: Category name e.g. 'Coverage Settings', 'Planned Orders',
                  'Production Order Lifecycle', 'Route & Operations', 'BOM Issues',
                  'Capacity Planning', 'Scheduling', 'Costing'
    """
    try:
        cases = sp.get_cases_by_filter(category=category)
        if not cases:
            return [{"error": f"No cases found for category: {category}"}]
        return [
            {
                "case_id":    c["case_id"],
                "symptom":    c["symptom"],
                "difficulty": c["difficulty"],
                "module":     c["module"],
            }
            for c in cases
        ]
    except Exception as e:
        log.error(f"get_cases_by_category error: {e}")
        return [{"error": str(e)}]


# ─────────────────────────────────────────────
# TOOL 4 — score_answer
# ─────────────────────────────────────────────

@mcp.tool()
def score_answer(case_id: str, user_answer: str) -> dict:
    """
    Return the correct answer data for a drill card so the
    calling model can evaluate the user's diagnosis.
    Also returns the Microsoft Learn reference URL for cross-validation.

    Args:
        case_id:     The case ID from get_random_drill e.g. 'PC-426B3412'
        user_answer: The user's diagnosis attempt as free text
    """
    try:
        case = sp.get_case_by_id(case_id)
        if not case:
            return {"error": f"Case not found: {case_id}"}

        return {
            "case_id":                  case["case_id"],
            "user_answer":              user_answer,
            "correct_root_cause":       case["root_cause"],
            "correct_diagnostic_steps": case["diagnostic_steps"],
            "correct_resolution":       case["resolution"],
            "d365_navigation":          case["d365_navigation"],
            "consultant_tip":           case["consultant_tip"],
            "learn_url":                case.get("learn_url", ""),
            "instruction": (
                "Compare the user_answer against correct_root_cause. "
                "Score 1-10. Show: what they got right, gaps, correct root cause, "
                "resolution, consultant tip. "
                "After scoring, ALWAYS call get_learn_reference with this case_id "
                "to cross-validate against Microsoft Learn official documentation. "
                "If learn_url is not empty, include it as: "
                "'✓ Confirmed by Microsoft Learn: [url]' or "
                "'⚠ Microsoft Learn also mentions: [additional detail]'. "
                "Format clearly with score, verdict, gaps, and learn reference."
            ),
        }
    except Exception as e:
        log.error(f"score_answer error: {e}")
        return {"error": str(e)}


# ─────────────────────────────────────────────
# TOOL 5 — get_weak_areas
# ─────────────────────────────────────────────

@mcp.tool()
def get_weak_areas(session_results: list[dict]) -> dict:
    """
    Analyse drill session history and return weak areas.
    Call this after every 5-10 drills to give the user performance feedback.

    Args:
        session_results: List of scoring results from score_answer calls.
                        Each item should have: case_id, score, category, module.
                        Example: [{"case_id": "PC-426B3412", "score": 6,
                                   "category": "Coverage Settings",
                                   "module": "Master Planning"}]
    """
    try:
        if not session_results:
            return {"error": "No session results provided."}

        from collections import defaultdict

        category_scores: dict = defaultdict(list)
        module_scores: dict = defaultdict(list)

        for r in session_results:
            score    = r.get("score", 0)
            category = r.get("category", "Unknown")
            module   = r.get("module", "Unknown")
            category_scores[category].append(score)
            module_scores[module].append(score)

        def avg(scores):
            return round(sum(scores) / len(scores), 1) if scores else 0

        category_avgs = {cat: avg(scores) for cat, scores in category_scores.items()}
        module_avgs   = {mod: avg(scores) for mod, scores in module_scores.items()}

        weak_categories = sorted(
            [(cat, sc) for cat, sc in category_avgs.items() if sc < 7],
            key=lambda x: x[1],
        )
        strong_categories = sorted(
            [(cat, sc) for cat, sc in category_avgs.items() if sc >= 7],
            key=lambda x: x[1], reverse=True,
        )

        overall = avg([r.get("score", 0) for r in session_results])

        recommendation = ""
        if weak_categories:
            worst = weak_categories[0][0]
            recommendation = (
                f"Focus your next session on '{worst}' — "
                f"it has your lowest average score of {weak_categories[0][1]}/10."
            )
        else:
            recommendation = "Strong performance across all categories. Try advanced difficulty next."

        return {
            "total_drills":        len(session_results),
            "overall_average":     overall,
            "category_scores":     category_avgs,
            "module_scores":       module_avgs,
            "weak_categories":     [{"category": c, "avg_score": s} for c, s in weak_categories],
            "strong_categories":   [{"category": c, "avg_score": s} for c, s in strong_categories],
            "recommendation":      recommendation,
        }

    except Exception as e:
        log.error(f"get_weak_areas error: {e}")
        return {"error": str(e)}


# ─────────────────────────────────────────────
# TOOL 6 — get_diagnostic_path
# ─────────────────────────────────────────────

@mcp.tool()
def get_diagnostic_path(case_id: str) -> dict:
    """
    Return the full step-by-step diagnostic path for a specific case.
    Use this after scoring to explain the correct investigation sequence.

    Args:
        case_id: The case ID e.g. 'PC-426B3412'
    """
    try:
        case = sp.get_case_by_id(case_id)
        if not case:
            return {"error": f"Case not found: {case_id}"}

        return {
            "case_id":          case["case_id"],
            "symptom":          case["symptom"],
            "root_cause":       case["root_cause"],
            "diagnostic_steps": case["diagnostic_steps"],
            "d365_navigation":  case["d365_navigation"],
            "resolution":       case["resolution"],
            "consultant_tip":   case["consultant_tip"],
            "source_url":       case["source_url"],
        }
    except Exception as e:
        log.error(f"get_diagnostic_path error: {e}")
        return {"error": str(e)}


# ─────────────────────────────────────────────
# TOOL 7 — get_odata_query
# ─────────────────────────────────────────────

@mcp.tool()
def get_odata_query(case_id: str) -> dict:
    """
    Return a D365 F&O OData query to verify the issue described in a drill card.
    Use this to show how to investigate the symptom directly in D365 data.

    Args:
        case_id: The case ID e.g. 'PC-426B3412'
    """
    try:
        case = sp.get_case_by_id(case_id)
        if not case:
            return {"error": f"Case not found: {case_id}"}

        return {
            "case_id":    case_id,
            "symptom":    case["symptom"],
            "module":     case["module"],
            "category":   case["category"],
            "root_cause": case["root_cause"],
            "instruction": (
                "Generate a practical D365 F&O OData query a consultant could run "
                "to investigate or verify this issue. "
                "Return: entity name (e.g. ProdTable, ReqTransCov), "
                "full OData query URL path (e.g. /data/ProdTable?$filter=...&$select=...), "
                "purpose (one sentence), and key_fields list. "
                "Use standard D365 F&O OData entities. Omit the base URL."
            ),
        }

    except Exception as e:
        log.error(f"get_odata_query error: {e}")
        return {"error": str(e)}




# ─────────────────────────────────────────────
# TOOL 8 — get_learn_reference
# ─────────────────────────────────────────────

@mcp.tool()
def get_learn_reference(case_id: str) -> dict:
    """
    Fetch the Microsoft Learn reference page for a drill card
    and return relevant sections for answer validation.

    Args:
        case_id: The case ID e.g. 'PC-426B3412'
    """
    try:
        case = sp.get_case_by_id(case_id)
        if not case:
            return {"error": f"Case not found: {case_id}"}

        learn_url = case.get("learn_url", "")
        if not learn_url or learn_url.endswith("..."):
            return {
                "case_id":  case_id,
                "status":   "no_learn_url",
                "message":  "No Microsoft Learn URL stored for this case.",
                "symptom":  case["symptom"],
            }

        # Fetch the Learn page
        import requests as req
        fetch_headers = {
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/120.0.0.0 Safari/537.36"
            )
        }
        resp = req.get(learn_url, headers=fetch_headers, timeout=15)

        if resp.status_code != 200:
            return {
                "case_id":   case_id,
                "status":    "fetch_failed",
                "learn_url": learn_url,
                "error":     f"HTTP {resp.status_code}",
            }

        from bs4 import BeautifulSoup
        soup = BeautifulSoup(resp.text, "html.parser")

        # Extract main article content
        article = soup.find("article") or soup.find("main") or soup.find("body")
        if not article:
            return {"error": "Could not parse Learn page content"}

        content_text = article.get_text(" ", strip=True)[:3000]

        return {
            "case_id":   case_id,
            "status":    "success",
            "learn_url": learn_url,
            "symptom":   case["symptom"],
            "content":   content_text,
            "instruction": (
                "Use this Microsoft Learn content to validate or supplement "
                "the forum-based answer. Highlight any additional details or "
                "corrections the official docs provide beyond the forum answer."
            ),
        }

    except Exception as e:
        log.error(f"get_learn_reference error: {e}")
        return {"error": str(e)}


# ─────────────────────────────────────────────
# TOOL 9 — start_auto_discover
# ─────────────────────────────────────────────

@mcp.tool()
def start_auto_discover(
    category: Optional[str] = None,
    limit: int = 20,
) -> dict:
    """
    Start a background job to discover and ingest new D365 forum threads.
    Returns immediately with a job_id. Use get_discover_status to check progress.
    Use when user asks to add more issues, expand the knowledge base,
    or find new drill cards on a specific topic.

    Args:
        category: Optional focus — 'Master Planning' or 'Production Control'.
                  If omitted, searches across all configured queries.
        limit:    Max new threads to process (default: 20).
    """
    import subprocess
    import sys
    from pathlib import Path

    auto_discover_script = Path(__file__).parent / "auto_discover.py"
    status_file = Path(__file__).parent / "discover_status.json"

    if not auto_discover_script.exists():
        return {"error": f"auto_discover.py not found at: {auto_discover_script}"}

    # Clear any previous status
    try:
        status_file.write_text(
            json.dumps({
                "stage": "starting",
                "detail": "Job queued",
                "updated_at": __import__("datetime").datetime.utcnow().isoformat(),
                "category": category or "all",
                "limit": limit,
            }),
            encoding="utf-8",
        )
    except Exception:
        pass

    cmd = [sys.executable, str(auto_discover_script), "--limit", str(limit)]
    if category:
        cmd += ["--category", category]

    log.info(f"start_auto_discover: launching background job: {' '.join(cmd)}")

    try:
        # Launch detached — does not block
        subprocess.Popen(
            cmd,
            cwd=str(Path(__file__).parent),
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            close_fds=True,
        )

        return {
            "status":   "started",
            "category": category or "all",
            "limit":    limit,
            "instruction": (
                "The discovery pipeline has been started in the background. "
                "Tell the user it is running and will take 10-15 minutes. "
                "They can ask 'what's the discovery status?' at any time "
                "and you will call get_discover_status to show live progress."
            ),
        }

    except Exception as e:
        log.error(f"start_auto_discover error: {e}")
        return {"error": str(e)}


# ─────────────────────────────────────────────
# TOOL 10 — get_discover_status
# ─────────────────────────────────────────────

@mcp.tool()
def get_discover_status() -> dict:
    """
    Check the current status of a running or completed auto-discovery job.
    Call this when the user asks about discovery progress or results.
    """
    from pathlib import Path

    status_file = Path(__file__).parent / "discover_status.json"

    if not status_file.exists():
        return {
            "status":  "no_job",
            "message": "No discovery job has been started yet.",
        }

    try:
        data = json.loads(status_file.read_text(encoding="utf-8"))
        stage = data.get("stage", "unknown")

        # Build a human-readable summary
        if stage == "starting":
            summary = "Job is queued and starting up..."
        elif stage == "searching":
            found = data.get("found_so_far", 0)
            summary = f"Searching forum queries... {found} new threads found so far."
        elif stage == "discovered":
            found = data.get("found", 0)
            summary = f"Discovery complete — {found} new threads found. Now scraping and enriching..."
        elif stage == "scraping":
            n = data.get("to_process", 0)
            summary = f"Scraping and enriching {n} threads via Claude API. This takes ~10 min..."
        elif stage == "pushing":
            summary = "Enrichment done. Pushing new cards to Dataverse..."
        elif stage == "complete":
            pushed  = data.get("pushed", 0)
            skipped = data.get("skipped", 0)
            errors  = data.get("errors", 0)
            summary = (
                f"Pipeline complete! {pushed} new cards added to Dataverse "
                f"({skipped} duplicates skipped, {errors} errors)."
            )
        else:
            summary = f"Stage: {stage}"

        return {
            "stage":      stage,
            "summary":    summary,
            "updated_at": data.get("updated_at", ""),
            "detail":     data.get("detail", ""),
            **{k: v for k, v in data.items()
               if k not in ("stage", "detail", "updated_at")},
        }

    except Exception as e:
        return {"error": f"Could not read status file: {e}"}


# ─────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────

if __name__ == "__main__":
    import sys

    # Claude Desktop launches via stdio (no tty)
    # Manual run uses HTTP for Copilot Studio / ngrok
    if sys.stdin.isatty() or "--http" in sys.argv:
        log.info("Starting D365 Drill Engine MCP Server on port 8002...")
        log.info("MCP endpoint: http://localhost:8002/mcp")
        log.info("Expose via ngrok: ngrok http 8002")
        mcp.run(transport="streamable-http", host="0.0.0.0", port=8002)
    else:
        mcp.run(transport="stdio")