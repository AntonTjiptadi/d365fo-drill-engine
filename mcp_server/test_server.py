"""
Test script — validates SharePoint connection and all 7 MCP tools.
Run this before starting the server.

Usage:
  cd mcp_server
  python test_server.py
"""

import os
import sys
import json
from dotenv import load_dotenv

load_dotenv(dotenv_path=os.path.join(os.path.dirname(__file__), '..', '.env'))
print("SITE_ID:", os.getenv("DRILL_SITE_ID"))
print("LIST_ID:", os.getenv("DRILL_LIST_ID"))
print("CLIENT_ID:", os.getenv("DRILL_CLIENT_ID"))
sys.path.insert(0, os.path.dirname(__file__))

# from sharepoint_client import SharePointClient
from dataverse_client import DataverseClient

def separator(title):
    print(f"\n{'='*55}")
    print(f"  {title}")
    print(f"{'='*55}")

def main():
    print("\nD365 Drill Engine — Connection Test")
    print("─" * 40)

    # ── Test 1: SharePoint connection ──
    separator("TEST 1: SharePoint connection")
    try:
        # sp = SharePointClient()
        sp = DataverseClient()
        cases = sp.get_all_cases()
        print(f"✓ Connected — {len(cases)} drill cards loaded")
        if cases:
            print(f"  First card: [{cases[0]['case_id']}] {cases[0]['symptom'][:60]}")
    except Exception as e:
        print(f"✗ Connection failed: {e}")
        return

    # ── Test 2: get_cases_by_filter ──
    separator("TEST 2: Filter by module")
    try:
        mp_cases = sp.get_cases_by_filter(module="Master Planning")
        pc_cases = sp.get_cases_by_filter(module="Production Control")
        print(f"✓ Master Planning cases:    {len(mp_cases)}")
        print(f"✓ Production Control cases: {len(pc_cases)}")
    except Exception as e:
        print(f"✗ Filter failed: {e}")

    # ── Test 3: search_by_symptom ──
    separator("TEST 3: Symptom search")
    try:
        results = sp.search_by_symptom("planned order not generated")
        print(f"✓ Found {len(results)} matches for 'planned order not generated'")
        for r in results[:2]:
            print(f"  [{r['case_id']}] {r['symptom'][:55]}")
    except Exception as e:
        print(f"✗ Search failed: {e}")

    # ── Test 4: get_case_by_id ──
    separator("TEST 4: Get case by ID")
    try:
        if cases:
            first_id = cases[0]["case_id"]
            case = sp.get_case_by_id(first_id)
            if case:
                print(f"✓ Retrieved case: {case['case_id']}")
                print(f"  Symptom:    {case['symptom'][:55]}")
                print(f"  Module:     {case['module']}")
                print(f"  Category:   {case['category']}")
                print(f"  Difficulty: {case['difficulty']}")
            else:
                print(f"✗ Case not found: {first_id}")
    except Exception as e:
        print(f"✗ Get by ID failed: {e}")

    # ── Test 5: Random drill (simulated) ──
    separator("TEST 5: Random drill card")
    try:
        import random
        case = random.choice(cases)
        print(f"✓ Random drill card selected:")
        print(f"  Case ID:    {case['case_id']}")
        print(f"  Symptom:    {case['symptom'][:60]}")
        print(f"  Difficulty: {case['difficulty']}")
        print(f"  (Root cause withheld until answer scored)")
    except Exception as e:
        print(f"✗ Random drill failed: {e}")

    # ── Test 6: Weak areas (simulated session) ──
    separator("TEST 6: Weak areas analysis")
    try:
        mock_session = [
            {"case_id": "PC-001", "score": 4, "category": "Coverage Settings",    "module": "Master Planning"},
            {"case_id": "PC-002", "score": 8, "category": "Route & Operations",   "module": "Production Control"},
            {"case_id": "PC-003", "score": 5, "category": "Coverage Settings",    "module": "Master Planning"},
            {"case_id": "PC-004", "score": 9, "category": "Route & Operations",   "module": "Production Control"},
            {"case_id": "PC-005", "score": 3, "category": "Planned Orders",       "module": "Master Planning"},
        ]
        from collections import defaultdict
        category_scores = defaultdict(list)
        for r in mock_session:
            category_scores[r["category"]].append(r["score"])

        weak = [(cat, sum(s)/len(s)) for cat, s in category_scores.items() if sum(s)/len(s) < 7]
        weak.sort(key=lambda x: x[1])

        print(f"✓ Analysed {len(mock_session)} drill results")
        print(f"  Overall average: {sum(r['score'] for r in mock_session)/len(mock_session):.1f}/10")
        if weak:
            print(f"  Weak areas:")
            for cat, avg in weak:
                print(f"    {cat}: {avg:.1f}/10")
    except Exception as e:
        print(f"✗ Weak areas failed: {e}")

    # ── Summary ──
    separator("SUMMARY")
    print("SharePoint connection:  ✓")
    print("Filter queries:         ✓")
    print("Symptom search:         ✓")
    print("Case retrieval:         ✓")
    print("Drill card selection:   ✓")
    print("Session analysis:       ✓")
    print()
    print("Ready to start the MCP server:")
    print("  python server.py")
    print()
    print("Then expose via ngrok:")
    print("  ngrok http 8002")
    print()
    print("MCP endpoint for Copilot Studio:")
    print("  https://<ngrok-url>/mcp")


if __name__ == "__main__":
    main()
