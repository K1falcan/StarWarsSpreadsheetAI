"""Thin client for the Star Wars API at https://swapi.info.

swapi.info returns every record for a resource in a single JSON array (there is
no pagination), and it cross-references other resources by absolute URL, e.g. a
person's ``homeworld`` is ``"https://swapi.info/api/planets/1"``. This module
fetches the six resource collections and exposes them as plain Python lists, plus
a URL->display-name index so those cross-references can be rendered as readable
text in the spreadsheet.
"""

from __future__ import annotations

import logging
from typing import Any, Dict, List

import requests
from requests.adapters import HTTPAdapter, Retry

log = logging.getLogger(__name__)

# Note the trailing-less base; swapi.info issues a 301 from http -> https, so we
# talk https directly to avoid an extra round trip on every call.
BASE_URL = "https://swapi.info/api"

# The six collections exposed at the API root.
RESOURCES = ("films", "people", "planets", "species", "starships", "vehicles")


class SwapiClient:
    """Fetches and caches the swapi.info resource collections."""

    def __init__(self, base_url: str = BASE_URL, timeout: int = 30) -> None:
        self.base_url = base_url.rstrip("/")
        self.timeout = timeout
        self.session = self._build_session()
        # Populated lazily by fetch_all(); keyed by resource name.
        self._collections: Dict[str, List[Dict[str, Any]]] = {}
        # Maps every resource "url" -> its human-readable label.
        self._name_by_url: Dict[str, str] = {}

    @staticmethod
    def _build_session() -> requests.Session:
        """A session that retries transient failures with backoff."""
        session = requests.Session()
        retry = Retry(
            total=4,
            backoff_factor=0.6,  # 0.6s, 1.2s, 2.4s, ...
            status_forcelist=(429, 500, 502, 503, 504),
            allowed_methods=("GET",),
        )
        adapter = HTTPAdapter(max_retries=retry)
        session.mount("https://", adapter)
        session.mount("http://", adapter)
        session.headers.update({"Accept": "application/json"})
        return session

    def _get(self, url: str) -> Any:
        log.debug("GET %s", url)
        response = self.session.get(url, timeout=self.timeout)
        response.raise_for_status()
        return response.json()

    def fetch_collection(self, resource: str) -> List[Dict[str, Any]]:
        """Fetch one resource collection (e.g. ``"people"``) as a list."""
        if resource not in RESOURCES:
            raise ValueError(f"Unknown resource {resource!r}; expected one of {RESOURCES}")
        url = f"{self.base_url}/{resource}"
        data = self._get(url)
        if not isinstance(data, list):
            raise TypeError(f"Expected a list from {url}, got {type(data).__name__}")
        log.info("Fetched %d %s", len(data), resource)
        return data

    def fetch_all(self) -> Dict[str, List[Dict[str, Any]]]:
        """Fetch all six collections and build the URL->name index."""
        self._collections = {res: self.fetch_collection(res) for res in RESOURCES}
        self._build_name_index()
        return self._collections

    def _build_name_index(self) -> None:
        """Index every record's ``url`` to a display label.

        People/planets/species/etc. expose ``name``; films expose ``title``.
        """
        index: Dict[str, str] = {}
        for records in self._collections.values():
            for record in records:
                url = record.get("url")
                if not url:
                    continue
                index[url] = record.get("name") or record.get("title") or url
        self._name_by_url = index
        log.info("Indexed %d resource URLs", len(index))

    def name_for(self, url: str) -> str:
        """Resolve a single resource URL to its label (or the URL if unknown)."""
        return self._name_by_url.get(url, url)

    def names_for(self, urls: List[str]) -> List[str]:
        """Resolve a list of resource URLs to their labels."""
        return [self.name_for(u) for u in urls]
