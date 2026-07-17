# ─────────────────────────────────────────────
# TOOL 4 — score_answer (updated with Learn cross-validation)
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
            "case_id":                   case["case_id"],
            "user_answer":               user_answer,
            "correct_root_cause":        case["root_cause"],
            "correct_diagnostic_steps":  case["diagnostic_steps"],
            "correct_resolution":        case["resolution"],
            "d365_navigation":           case["d365_navigation"],
            "consultant_tip":            case["consultant_tip"],
            "learn_url":                 case.get("learn_url", ""),
            "instruction": (
                "Compare the user_answer against correct_root_cause. "
                "Score 1-10. Show: what they got right, gaps, correct root cause, "
                "resolution, consultant tip. "
                "If learn_url is not empty, include it as '✓ Confirmed by Microsoft Learn: [url]' "
                "or '⚠ Microsoft Learn also mentions: [detail from url context]'. "
                "Format the response clearly with score, verdict, gaps, and learn reference."
            ),
        }
    except Exception as e:
        log.error(f"score_answer error: {e}")
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
        if not learn_url:
            return {
                "case_id":  case_id,
                "status":   "no_learn_url",
                "message":  "No Microsoft Learn URL stored for this case.",
                "symptom":  case["symptom"],
            }

        # Fetch the Learn page
        import requests as req
        headers = {
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/120.0.0.0 Safari/537.36"
            )
        }
        resp = req.get(learn_url, headers=headers, timeout=15)

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

        # Get text — limit to 3000 chars to avoid overwhelming context
        content = article.get_text(" ", strip=True)[:3000]

        return {
            "case_id":   case_id,
            "status":    "success",
            "learn_url": learn_url,
            "symptom":   case["symptom"],
            "content":   content,
            "instruction": (
                "Use this Microsoft Learn content to validate or supplement "
                "the forum-based answer. Highlight any additional details or "
                "corrections the official docs provide beyond the forum answer."
            ),
        }

    except Exception as e:
        log.error(f"get_learn_reference error: {e}")
        return {"error": str(e)}
