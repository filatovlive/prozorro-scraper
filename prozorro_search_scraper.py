 #!/usr/bin/env python3
"""
Prozorro scraper with filters configured directly in this file.

The script uses the public search API in a UI-like mode and writes the result
to an Excel file.
"""

from __future__ import annotations

import json
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import dataclass
from math import ceil
from pathlib import Path
from typing import Any, Dict, List, Sequence, Tuple

from openpyxl import Workbook


ORG_URL = "https://prozorro.gov.ua/api/search/organizations"
SEARCH_TENDERS_URL = "https://prozorro.gov.ua/api/search/tenders"
SUMMARY_TENDER_URL = "https://prozorro.gov.ua/api/tenders/{tender_id}/summary"
DEFAULT_TIMEOUT = 60
RETRYABLE_HTTP_CODES = {429, 502, 503, 504}
MAX_RETRIES = 7
VALID_ENTITIES = {"tenders"}
ORG_FILTER_TYPES = {
    "buyer": "buyer",
    "tenderer": "tenderer",
    "supplier": "supplier",
    "producer": "productVendor",
    "procuring-entity": "procuringEntity",
}
SCRIPT_DIR = Path(__file__).resolve().parent


# Edit filters here.
SCRAPER_CONFIG: Dict[str, Any] = {
    "entity": "tenders",
    "timeout": 60,
    "resolve_org_names": True,
    "resolve_exact_name_first": True,
    "fetch_all_pages": True,
    "max_pages": 300,
    "per_page": 100,
    "enrich_with_summary_id": True,
    "summary_request_delay": 0.5,
    "output_path": SCRIPT_DIR / "output" / "prozorro_tenders_firm.xlsx",
    "filters": {
        "text": "",
        "buyer": ['ТОВ "СУПЕРСИМЕТРІЯ"'],
        "tenderer": [],
        "supplier": [],
        "producer": [],
        "procuringEntity": [],
        "status": [],
        "proc_type": [],
        "cpv": [],
        "cpv_mask": [],
        "region": "",
        "year": [2026],
        "funders": [],
        "awardCriteria": [],
        "proc_rationale": [],
        "milestone": None,
        "local_share": "",
        "yearCreated": [],
        "contract_status": [],
        "product_status": [],
        "framework_types": [],
        "agreement": [],
        "group": "",
        "sort_by": "",
        "order": "",
        "page": None,
        "per_page": None,
        "value": {
            "currency": "UAH",
            "amount": {"start": "", "end": ""},
        },
        "contract_value": {
            "currency": "UAH",
            "amount": {"start": "", "end": ""},
        },
        "date": {
            "tender": {"start": "", "end": ""},
            "enquiry": {"start": "", "end": ""},
            "auction": {"start": "", "end": ""},
            "award": {"start": "", "end": ""},
            "plan": {"start": "", "end": ""},
            "dateSigned": {"start": "", "end": ""},
            "framework": {"start": "", "end": ""},
            "qualification": {"start": "", "end": ""},
        },
    },
}


class ProzorroSearchError(RuntimeError):
    """Raised when the Prozorro search API returns an error."""


@dataclass
class SearchRunMeta:
    mode: str
    text_queries: List[str]
    total_by_query: Dict[str, int]
    pages_fetched_by_query: Dict[str, int]
    summary_id_matches: int = 0
    summary_requests_done: int = 0


def item_identity(item: Dict[str, Any]) -> str:
    for key in ("tenderID", "contractID", "planID", "prettyID", "id"):
        value = item.get(key)
        if value:
            return f"{key}:{value}"
    return json.dumps(item, ensure_ascii=False, sort_keys=True)


def _build_query_pairs(params: Dict[str, Any], prefix: str = "") -> List[Tuple[str, str]]:
    pairs: List[Tuple[str, str]] = []
    for key, value in params.items():
        if value is None:
            continue
        full_key = f"{prefix}.{key}" if prefix else key
        if isinstance(value, dict):
            pairs.extend(_build_query_pairs(value, full_key))
        elif isinstance(value, (list, tuple)):
            for item in value:
                if item is None:
                    continue
                if isinstance(item, dict):
                    pairs.extend(_build_query_pairs(item, full_key))
                else:
                    pairs.append((full_key, str(item)))
        else:
            pairs.append((full_key, str(value)))
    return pairs


def _post_with_query(url: str, params: Dict[str, Any], timeout: int) -> Dict[str, Any]:
    query = urllib.parse.urlencode(_build_query_pairs(params), doseq=True)
    full_url = f"{url}?{query}" if query else url
    for attempt in range(MAX_RETRIES + 1):
        request = urllib.request.Request(
            full_url,
            data=b"",
            method="POST",
            headers={
                "Accept": "application/json",
                "Content-Type": "application/json",
                "User-Agent": "prozorro-search-scraper/1.0",
            },
        )
        try:
            with urllib.request.urlopen(request, timeout=timeout) as response:
                return json.loads(response.read().decode("utf-8"))
        except urllib.error.HTTPError as exc:
            body = exc.read().decode("utf-8", errors="replace")
            if exc.code in RETRYABLE_HTTP_CODES and attempt < MAX_RETRIES:
                time.sleep(2 ** attempt)
                continue
            raise ProzorroSearchError(f"HTTP {exc.code} for {full_url}: {body}") from exc
        except urllib.error.URLError as exc:
            raise ProzorroSearchError(f"Request failed for {full_url}: {exc.reason}") from exc
        except json.JSONDecodeError as exc:
            raise ProzorroSearchError(f"Invalid JSON response from {full_url}") from exc
    raise ProzorroSearchError(f"Request failed after retries for {full_url}")


def _get_with_query(url: str, params: Dict[str, Any], timeout: int) -> Dict[str, Any]:
    query = urllib.parse.urlencode(_build_query_pairs(params), doseq=True)
    full_url = f"{url}?{query}" if query else url
    return _get_json(full_url, timeout)


def _get_json(url: str, timeout: int) -> Dict[str, Any]:
    for attempt in range(MAX_RETRIES + 1):
        request = urllib.request.Request(
            url,
            method="GET",
            headers={
                "Accept": "application/json",
                "User-Agent": "prozorro-search-scraper/1.0",
            },
        )
        try:
            with urllib.request.urlopen(request, timeout=timeout) as response:
                return json.loads(response.read().decode("utf-8"))
        except urllib.error.HTTPError as exc:
            body = exc.read().decode("utf-8", errors="replace")
            if exc.code in RETRYABLE_HTTP_CODES and attempt < MAX_RETRIES:
                time.sleep(2 ** attempt)
                continue
            raise ProzorroSearchError(f"HTTP {exc.code} for {url}: {body}") from exc
        except urllib.error.URLError as exc:
            raise ProzorroSearchError(f"Request failed for {url}: {exc.reason}") from exc
        except json.JSONDecodeError as exc:
            raise ProzorroSearchError(f"Invalid JSON response from {url}") from exc
    raise ProzorroSearchError(f"Request failed after retries for {url}")


def search_organizations(search_type: str, value: str, timeout: int) -> List[Dict[str, Any]]:
    payload = _post_with_query(ORG_URL, {"type": search_type, "value": value}, timeout)
    items = payload.get("data")
    if not isinstance(items, list):
        raise ProzorroSearchError("Unexpected organization lookup response shape")
    return items


def resolve_org_values(kind: str, values: Sequence[str], exact_first: bool, timeout: int) -> List[str]:
    resolved: List[str] = []
    api_type = ORG_FILTER_TYPES[kind]
    for raw in values:
        candidate = str(raw).strip()
        if not candidate:
            continue
        if candidate.isdigit():
            resolved.append(candidate)
            continue
        items = search_organizations(api_type, candidate, timeout)
        if not items:
            raise ProzorroSearchError(f'No organizations found for {kind} value "{candidate}"')
        if exact_first:
            exact_matches = [item for item in items if str(item.get("name", "")).casefold() == candidate.casefold()]
            picked = exact_matches[0] if exact_matches else items[0]
        else:
            picked = items[0]
        org_id = str(picked.get("id", "")).strip()
        if not org_id:
            raise ProzorroSearchError(f'Organization lookup for "{candidate}" returned empty id')
        resolved.append(org_id)
    return resolved


def _clean_value(value: Any) -> Any:
    if isinstance(value, dict):
        cleaned = {k: _clean_value(v) for k, v in value.items()}
        return {k: v for k, v in cleaned.items() if v not in (None, "", [], {})}
    if isinstance(value, list):
        cleaned = [_clean_value(item) for item in value]
        return [item for item in cleaned if item not in (None, "", [], {})]
    return value


def build_params_from_config(config: Dict[str, Any]) -> Dict[str, Any]:
    params = _clean_value(config["filters"])

    milestone = params.get("milestone")
    if isinstance(milestone, str):
        params["milestone"] = {"code": milestone}

    if not params.get("value", {}).get("amount"):
        params.pop("value", None)
    if not params.get("contract_value", {}).get("amount"):
        params.pop("contract_value", None)

    return params


def ensure_supported_ui_filters(params: Dict[str, Any]) -> None:
    supported_keys = {"buyer", "text", "year"}
    unsupported_keys = sorted(key for key, value in params.items() if key not in supported_keys and value not in (None, "", [], {}))
    if unsupported_keys:
        raise ProzorroSearchError(
            "UI-like mode currently supports only `buyer`, `text`, and local `year`. Unsupported non-empty filters: "
            + ", ".join(unsupported_keys)
        )


def maybe_resolve_org_filters(params: Dict[str, Any], config: Dict[str, Any]) -> Dict[str, Any]:
    if not config.get("resolve_org_names", True):
        return params

    mapping = {
        "buyer": "buyer",
        "tenderer": "tenderer",
        "supplier": "supplier",
        "vendor": "producer",
        "procuringEntity": "procuring-entity",
    }
    resolved = dict(params)
    for api_key, kind in mapping.items():
        values = resolved.get(api_key)
        if not values:
            continue
        if not isinstance(values, list):
            values = [values]
        resolved[api_key] = resolve_org_values(
            kind=kind,
            values=values,
            exact_first=config.get("resolve_exact_name_first", True),
            timeout=int(config.get("timeout", DEFAULT_TIMEOUT)),
        )
    return resolved


def fetch_search_tenders(params: Dict[str, Any], timeout: int) -> Dict[str, Any]:
    payload = _post_with_query(SEARCH_TENDERS_URL, params, timeout)
    items = payload.get("data")
    if not isinstance(items, list):
        raise ProzorroSearchError("Unexpected search response shape: `data` is not a list")
    return payload


def enrich_items_with_summary_ids(
    items: Sequence[Dict[str, Any]],
    timeout: int,
    request_delay: float,
) -> Tuple[List[Dict[str, Any]], int, int]:
    enriched: List[Dict[str, Any]] = []
    matches = 0
    requests_done = 0
    for item in items:
        clone = dict(item)
        tender_id = str(clone.get("tenderID", "")).strip()
        if tender_id:
            payload = _get_json(SUMMARY_TENDER_URL.format(tender_id=urllib.parse.quote(tender_id, safe="")), timeout)
            requests_done += 1
            object_id = str(payload.get("id", "")).strip()
            if object_id:
                clone["id"] = object_id
                matches += 1
            if request_delay > 0:
                time.sleep(request_delay)
        enriched.append(clone)
    return enriched, matches, requests_done


def deduplicate_items(items: Sequence[Dict[str, Any]]) -> List[Dict[str, Any]]:
    unique_items: List[Dict[str, Any]] = []
    seen_ids: set[str] = set()
    for item in items:
        identity = item_identity(item)
        if identity in seen_ids:
            continue
        seen_ids.add(identity)
        unique_items.append(item)
    return unique_items


def item_year(item: Dict[str, Any]) -> str:
    tender_id = str(item.get("tenderID", "")).strip()
    if tender_id.startswith("UA-") and len(tender_id) >= 7:
        return tender_id[3:7]
    for key in ("dateCreated", "dateModified", "date"):
        value = str(item.get(key, "")).strip()
        if len(value) >= 4 and value[:4].isdigit():
            return value[:4]
    return ""


def item_procuring_entity_id(item: Dict[str, Any]) -> str:
    procuring_entity = item.get("procuringEntity")
    if not isinstance(procuring_entity, dict):
        return ""
    identifier = procuring_entity.get("identifier")
    if not isinstance(identifier, dict):
        return ""
    return str(identifier.get("id", "")).strip()


def fetch_tenders_with_ui_like_text_search(
    buyer_ids: Sequence[str],
    text_value: str,
    year_values: Sequence[Any],
    timeout: int,
    fetch_all_pages: bool,
    max_pages: int,
    per_page: int,
) -> Tuple[List[Dict[str, Any]], SearchRunMeta]:
    all_items: List[Dict[str, Any]] = []
    total_by_query: Dict[str, int] = {}
    pages_fetched_by_query: Dict[str, int] = {}
    queries: List[str] = []
    target_years = {str(value).strip() for value in year_values if str(value).strip()}
    text_candidate = str(text_value).strip()
    if text_candidate:
        queries.append(text_candidate)
    else:
        for buyer_id in buyer_ids:
            query = str(buyer_id).strip()
            if query:
                queries.append(query)
    for query in queries:
        pages_fetched = 0
        for page in range(1, max_pages + 1):
            payload = fetch_search_tenders({"text": query, "page": page, "per_page": per_page}, timeout)
            if page == 1:
                total_by_query[query] = int(payload.get("total", 0))
            data = payload.get("data", [])
            if not isinstance(data, list) or not data:
                break
            pages_fetched += 1
            if target_years:
                data = [item for item in data if item_year(item) in target_years]
            if buyer_ids and not text_candidate:
                target_buyer_ids = {str(value).strip() for value in buyer_ids if str(value).strip()}
                data = [item for item in data if item_procuring_entity_id(item) in target_buyer_ids]
            all_items.extend(data)
            if not fetch_all_pages:
                break
            actual_per_page = int(payload.get("per_page", 0) or len(payload.get("data", [])) or per_page)
            total_items = int(payload.get("total", 0))
            if actual_per_page <= 0:
                break
            total_pages = ceil(total_items / actual_per_page) if total_items > 0 else 0
            if page >= total_pages:
                break
        pages_fetched_by_query[query] = pages_fetched
    items = deduplicate_items(all_items)
    run_meta = SearchRunMeta(
        mode="ui_like_text_search",
        text_queries=queries,
        total_by_query=total_by_query,
        pages_fetched_by_query=pages_fetched_by_query,
    )
    return items, run_meta


def flatten_row(value: Any, prefix: str = "") -> Dict[str, Any]:
    flat: Dict[str, Any] = {}
    if isinstance(value, dict):
        for key, nested in value.items():
            nested_prefix = f"{prefix}.{key}" if prefix else key
            flat.update(flatten_row(nested, nested_prefix))
        return flat
    if isinstance(value, list):
        if all(not isinstance(item, (dict, list)) for item in value):
            flat[prefix] = "; ".join("" if item is None else str(item) for item in value)
            return flat
        for index, nested in enumerate(value):
            nested_prefix = f"{prefix}[{index}]"
            flat.update(flatten_row(nested, nested_prefix))
        return flat
    flat[prefix] = value
    return flat


def write_excel(
    output_path: Path,
    items: Sequence[Dict[str, Any]],
    entity: str,
    params: Dict[str, Any],
    run_meta: SearchRunMeta | None = None,
) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()

    data_sheet = workbook.active
    data_sheet.title = "data"

    flattened_rows = [flatten_row(item) for item in items]
    headers = sorted({key for row in flattened_rows for key in row.keys()})
    if headers:
        data_sheet.append(headers)
        for row in flattened_rows:
            data_sheet.append([row.get(header, "") for header in headers])
    else:
        data_sheet.append(["message"])
        data_sheet.append(["No data returned"])

    meta_sheet = workbook.create_sheet("meta")
    meta_sheet.append(["entity", entity])
    meta_sheet.append(["total_items_returned", len(items)])
    meta_sheet.append(["params_json", json.dumps(params, ensure_ascii=False)])
    if run_meta is not None:
        meta_sheet.append(["mode", run_meta.mode])
        meta_sheet.append(["text_queries_json", json.dumps(run_meta.text_queries, ensure_ascii=False)])
        meta_sheet.append(["total_by_query_json", json.dumps(run_meta.total_by_query, ensure_ascii=False)])
        meta_sheet.append(["pages_fetched_by_query_json", json.dumps(run_meta.pages_fetched_by_query, ensure_ascii=False)])
        meta_sheet.append(["summary_id_matches", run_meta.summary_id_matches])
        meta_sheet.append(["summary_requests_done", run_meta.summary_requests_done])

    workbook.save(output_path)


def validate_config(config: Dict[str, Any]) -> None:
    entity = config.get("entity")
    if entity not in VALID_ENTITIES:
        raise ProzorroSearchError(f"Invalid entity in SCRAPER_CONFIG: {entity}")
    output_path = config.get("output_path")
    if not output_path:
        raise ProzorroSearchError("output_path is required in SCRAPER_CONFIG")
    if Path(output_path).suffix.lower() != ".xlsx":
        raise ProzorroSearchError("output_path must point to an .xlsx file")


def main() -> int:
    try:
        validate_config(SCRAPER_CONFIG)
        params = build_params_from_config(SCRAPER_CONFIG)
        params = maybe_resolve_org_filters(params, SCRAPER_CONFIG)
        ensure_supported_ui_filters(params)

        entity = str(SCRAPER_CONFIG["entity"])
        timeout = int(SCRAPER_CONFIG.get("timeout", DEFAULT_TIMEOUT))
        output_path = Path(SCRAPER_CONFIG["output_path"])
        buyer_ids = [str(value) for value in params.get("buyer", [])]
        text_value = str(params.get("text", "")).strip()
        year_values = params.get("year", [])
        if not isinstance(year_values, list):
            year_values = [year_values]
        items, ui_meta = fetch_tenders_with_ui_like_text_search(
            buyer_ids,
            text_value,
            year_values,
            timeout,
            bool(SCRAPER_CONFIG.get("fetch_all_pages", False)),
            int(SCRAPER_CONFIG.get("max_pages", 1)),
            int(SCRAPER_CONFIG.get("per_page", 20)),
        )
        if SCRAPER_CONFIG.get("enrich_with_summary_id", True):
            items, summary_id_matches, summary_requests_done = enrich_items_with_summary_ids(
                items,
                timeout,
                float(SCRAPER_CONFIG.get("summary_request_delay", 0)),
            )
            ui_meta.summary_id_matches = summary_id_matches
            ui_meta.summary_requests_done = summary_requests_done
        write_excel(
            output_path,
            items,
            entity,
            params,
            ui_meta,
        )
        return 0
    except ProzorroSearchError as exc:
        sys.stderr.write(f"ERROR: {exc}\n")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
