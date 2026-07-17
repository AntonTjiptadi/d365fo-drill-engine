"""
Dataverse OData client for D365 Drill Engine.
Replaces sharepoint_client.py with clean OData queries.
Same interface — server.py only needs one import line changed.
"""

import os
import random
import logging
from typing import Optional
from dotenv import load_dotenv
import msal
import requests

load_dotenv(dotenv_path=os.path.join(os.path.dirname(__file__), '..', '.env'))

log = logging.getLogger(__name__)

ENTITY = "cdr_drillcards"


class DataverseClient:
    def __init__(self):
        self.dataverse_url = os.getenv("DATAVERSE_URL")
        self.tenant_id     = os.getenv("DRILL_TENANT_ID")
        self.client_id     = os.getenv("DRILL_CLIENT_ID")
        self.client_secret = os.getenv("DRILL_CLIENT_SECRET")
        self.base_url      = f"{self.dataverse_url}/api/data/v9.2"
        self.scope         = f"{self.dataverse_url}/.default"

        self._app = msal.ConfidentialClientApplication(
            self.client_id,
            authority=f"https://login.microsoftonline.com/{self.tenant_id}",
            client_credential=self.client_secret,
        )

    def _get_token(self) -> str:
        result = self._app.acquire_token_for_client(scopes=[self.scope])
        if "access_token" not in result:
            raise RuntimeError(f"Token error: {result.get('error_description')}")
        return result["access_token"]

    def _headers(self) -> dict:
        return {
            "Authorization":    f"Bearer {self._get_token()}",
            "OData-MaxVersion": "4.0",
            "OData-Version":    "4.0",
            "Accept":           "application/json",
        }

    def _select(self) -> str:
        """Standard field select — all lowercase logical names."""
        return (
            "cdr_symptom,cdr_caseid,cdr_casedate,cdr_module,cdr_category,"
            "cdr_rootcause,cdr_diagnosticsteps,cdr_d365navigation,"
            "cdr_resolution,cdr_difficulty,cdr_consultanttip,"
            "cdr_relatedconcepts,cdr_qualityscore,cdr_answerismsft,"
            "cdr_sourceurl,cdr_learnurl,cdr_rawtitle,cdr_scrapedat,cdr_drillcardid"
        )

    def get_all_cases(self) -> list[dict]:
        """Fetch all drill cards."""
        url = (
            f"{self.base_url}/{ENTITY}"
            f"?$select={self._select()}"
            f"&$top=1000"
        )
        resp = requests.get(url, headers=self._headers(), timeout=15)
        resp.raise_for_status()
        return [self._normalise(r) for r in resp.json().get("value", [])]

    def get_cases_by_filter(
        self,
        module: Optional[str] = None,
        category: Optional[str] = None,
        difficulty: Optional[str] = None,
    ) -> list[dict]:
        """Fetch drill cards with server-side OData filtering."""
        filters = []
        if module:
            filters.append(f"cdr_module eq '{module}'")
        if category:
            filters.append(f"cdr_category eq '{category}'")
        if difficulty:
            filters.append(f"cdr_difficulty eq '{difficulty}'")

        url = (
            f"{self.base_url}/{ENTITY}"
            f"?$select={self._select()}"
            f"&$top=1000"
        )
        if filters:
            url += "&$filter=" + " and ".join(filters)

        resp = requests.get(url, headers=self._headers(), timeout=15)
        resp.raise_for_status()
        return [self._normalise(r) for r in resp.json().get("value", [])]

    def get_case_by_id(self, case_id: str) -> Optional[dict]:
        """Fetch a single drill card by CaseID."""
        url = (
            f"{self.base_url}/{ENTITY}"
            f"?$select={self._select()}"
            f"&$filter=cdr_caseid eq '{case_id}'"
            f"&$top=1"
        )
        resp = requests.get(url, headers=self._headers(), timeout=15)
        resp.raise_for_status()
        items = resp.json().get("value", [])
        return self._normalise(items[0]) if items else None

    def search_by_symptom(self, symptom_text: str) -> list[dict]:
        """
        Keyword search across symptom and raw title.
        Uses OData contains filter — server-side search.
        """
        keyword = symptom_text.split()[0] if symptom_text else ""
        url = (
            f"{self.base_url}/{ENTITY}"
            f"?$select={self._select()}"
            f"&$filter=contains(cdr_symptom,'{keyword}') "
            f"or contains(cdr_rawtitle,'{keyword}')"
            f"&$top=5"
        )
        resp = requests.get(url, headers=self._headers(), timeout=15)
        resp.raise_for_status()
        results = [self._normalise(r) for r in resp.json().get("value", [])]

        # Re-score client side for better ranking
        keywords = symptom_text.lower().split()
        scored = []
        for case in results:
            text = (case.get("symptom", "") + " " + case.get("raw_title", "")).lower()
            score = sum(1 for kw in keywords if kw in text)
            scored.append((score, case))
        scored.sort(key=lambda x: x[0], reverse=True)
        return [c for _, c in scored]

    def get_stats(self) -> dict:
        """Return knowledge base statistics."""
        all_cases = self.get_all_cases()
        from collections import Counter
        modules      = Counter(c.get("module", "Unknown") for c in all_cases)
        categories   = Counter(c.get("category", "Unknown") for c in all_cases)
        difficulties = Counter(c.get("difficulty", "Unknown") for c in all_cases)
        return {
            "total":         len(all_cases),
            "by_module":     dict(modules),
            "by_category":   dict(categories),
            "by_difficulty": dict(difficulties),
        }

    def _normalise(self, record: dict) -> dict:
        """Map Dataverse logical names to clean dict."""
        answer_is_msft = record.get("cdr_answerismsft", False)
        if isinstance(answer_is_msft, str):
            answer_is_msft = answer_is_msft.lower() in ("yes", "true", "1")

        return {
            "case_id":          record.get("cdr_caseid", ""),
            "case_date":        record.get("cdr_casedate", ""),
            "symptom":          record.get("cdr_symptom", ""),
            "module":           record.get("cdr_module", ""),
            "category":         record.get("cdr_category", ""),
            "root_cause":       record.get("cdr_rootcause", ""),
            "diagnostic_steps": record.get("cdr_diagnosticsteps", ""),
            "d365_navigation":  record.get("cdr_d365navigation", ""),
            "resolution":       record.get("cdr_resolution", ""),
            "difficulty":       record.get("cdr_difficulty", ""),
            "consultant_tip":   record.get("cdr_consultanttip", ""),
            "related_concepts": record.get("cdr_relatedconcepts", ""),
            "quality_score":    record.get("cdr_qualityscore", 0),
            "answer_is_msft":   answer_is_msft,
            "source_url":       record.get("cdr_sourceurl", ""),
            "learn_url":        record.get("cdr_learnurl", ""),
            "raw_title":        record.get("cdr_rawtitle", ""),
            "scraped_at":       record.get("cdr_scrapedat", ""),
        }
