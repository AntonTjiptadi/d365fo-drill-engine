"""
SharePoint Graph API client for D365 Drill Engine.
Handles token acquisition and list queries via Microsoft Graph.
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

GRAPH_BASE = "https://graph.microsoft.com/v1.0"


class SharePointClient:
    def __init__(self):
        self.tenant_id     = os.getenv("DRILL_TENANT_ID")
        self.client_id     = os.getenv("DRILL_CLIENT_ID")
        self.client_secret = os.getenv("DRILL_CLIENT_SECRET")
        self.site_id       = os.getenv("DRILL_SITE_ID")
        self.list_id       = os.getenv("DRILL_LIST_ID")

        self._app = msal.ConfidentialClientApplication(
            self.client_id,
            authority=f"https://login.microsoftonline.com/{self.tenant_id}",
            client_credential=self.client_secret,
        )
        self._token: Optional[str] = None

    def _get_token(self) -> str:
        result = self._app.acquire_token_for_client(
            scopes=["https://graph.microsoft.com/.default"]
        )
        if "access_token" not in result:
            raise RuntimeError(f"Token error: {result.get('error_description')}")
        return result["access_token"]

    def _headers(self) -> dict:
        return {"Authorization": f"Bearer {self._get_token()}"}

    def _list_url(self) -> str:
        return f"{GRAPH_BASE}/sites/{self.site_id}/lists/{self.list_id}/items"

    def get_all_cases(self) -> list[dict]:
        url = f"{self._list_url()}?expand=fields&$top=999"
        resp = requests.get(url, headers=self._headers(), timeout=15)
        resp.raise_for_status()
        items = resp.json().get("value", [])
        return [self._normalise(i["fields"]) for i in items]

    def get_cases_by_filter(
    self,
    module: Optional[str] = None,
    category: Optional[str] = None,
    difficulty: Optional[str] = None,
    ) -> list[dict]:
        all_cases = self.get_all_cases()
        results = all_cases
        if module:
            results = [c for c in results if c.get("module", "").lower() == module.lower()]
        if category:
            results = [c for c in results if c.get("category", "").lower() == category.lower()]
        if difficulty:
            results = [c for c in results if c.get("difficulty", "").lower() == difficulty.lower()]
        return results

    def get_case_by_id(self, case_id: str) -> Optional[dict]:
        """Fetch a single drill card by CaseID field."""
        all_cases = self.get_all_cases()
        for case in all_cases:
            if case.get("case_id") == case_id:
                return case
        return None

    def search_by_symptom(self, symptom_text: str) -> list[dict]:
        """
        Basic keyword search across Title (symptom) field.
        Returns up to 5 closest matches.
        """
        all_cases = self.get_all_cases()
        keywords = symptom_text.lower().split()
        scored = []
        for case in all_cases:
            text = (case.get("symptom", "") + " " + case.get("raw_title", "")).lower()
            score = sum(1 for kw in keywords if kw in text)
            if score > 0:
                scored.append((score, case))
        scored.sort(key=lambda x: x[0], reverse=True)
        return [c for _, c in scored[:5]]

    def _normalise(self, fields: dict) -> dict:
        """Map SharePoint field names to clean dict."""
        return {
            "case_id":          fields.get("field_1", ""),
            "symptom":          fields.get("Title", ""),
            "module":           fields.get("field_2", ""),
            "category":         fields.get("field_3", ""),
            "root_cause":       fields.get("field_4", ""),
            "diagnostic_steps": fields.get("field_5", ""),
            "d365_navigation":  fields.get("field_6", ""),
            "resolution":       fields.get("field_7", ""),
            "difficulty":       fields.get("field_8", ""),
            "consultant_tip":   fields.get("field_9", ""),
            "related_concepts": fields.get("field_10", ""),
            "quality_score":    fields.get("field_11", 0),
            "answer_is_msft":   fields.get("field_12", "No"),
            "source_url":       fields.get("field_13", ""),
            "raw_title":        fields.get("field_14", ""),
        }