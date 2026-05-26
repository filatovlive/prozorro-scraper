import csv
import json
import re
import time
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import requests


# =============================================================================
# НАЛАШТУВАННЯ
# =============================================================================

# CSV-файл має містити колонки:
#   id        — внутрішній Prozorro ID тендера
#   tenderID  — публічний номер, наприклад UA-2026-01-16-001966-a
SCRIPT_DIR = Path(__file__).resolve().parent
INPUT_CSV = SCRIPT_DIR / "output" / "aem_2026.csv"

# Коренева папка для збереження JSON та документів
DOWNLOAD_DIR = SCRIPT_DIR / "downloaded_documents"

# Prozorro API
PROZORRO_API_BASE = "https://public-api.prozorro.gov.ua/api/2.5"

# Мережеві параметри
REQUEST_TIMEOUT_TENDER = 30
REQUEST_TIMEOUT_FILE = 90
MAX_RETRIES = 3
RETRY_SLEEP_SECONDS = 2

HEADERS = {
    "User-Agent": "prozorro-risk-analyzer/1.0 (+local-script)",
    "Accept": "application/json, */*",
}

FATAL_NETWORK_PREFIX = "FATAL_NETWORK:"
PUBLIC_TENDER_ID_RE = re.compile(r"UA-\d{4}-\d{2}-\d{2}-\d{6}-[a-z]")
INTERNAL_ID_RE = re.compile(r"\b[0-9a-f]{32}\b")


# =============================================================================
# ДОПОМІЖНІ ФУНКЦІЇ
# =============================================================================

def safe_filename(name: str, max_len: int = 180) -> str:
    """
    Робить назву безпечною для Windows/Linux/macOS.
    """
    if not name:
        return "unnamed"

    name = str(name).strip()
    name = re.sub(r'[<>:"/\\|?*\x00-\x1F]', "_", name)
    name = re.sub(r"\s+", " ", name)
    name = name.strip(" .")

    if not name:
        name = "unnamed"

    return name[:max_len]


def ensure_unique_path(path: Path) -> Path:
    """
    Якщо файл уже існує, додає суфікс _001, _002, ...
    """
    if not path.exists():
        return path

    stem = path.stem
    suffix = path.suffix

    for i in range(1, 10000):
        candidate = path.with_name(f"{stem}_{i:03d}{suffix}")
        if not candidate.exists():
            return candidate

    raise RuntimeError(f"Не вдалося створити унікальне ім'я для файлу: {path}")


def request_with_retries(
    url: str,
    *,
    timeout: int,
    expect_json: bool = False,
) -> Tuple[Optional[requests.Response], Optional[str]]:
    """
    GET-запит з повторними спробами.
    Повертає (response, error).
    """
    last_error = None

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            response = requests.get(url, headers=HEADERS, timeout=timeout)

            if response.status_code == 200:
                if expect_json:
                    try:
                        response.json()
                    except Exception as e:
                        return None, f"Некоректний JSON: {e}"
                return response, None

            last_error = f"HTTP {response.status_code}"

        except requests.RequestException as e:
            last_error = str(e)
            error_details = f"{e} {e!r}".lower()
            if (
                "getaddrinfo failed" in error_details
                or "nameresolutionerror" in error_details
                or "failed to resolve" in error_details
            ):
                return None, f"{FATAL_NETWORK_PREFIX} DNS resolution failed: {last_error}"

        if attempt < MAX_RETRIES:
            time.sleep(RETRY_SLEEP_SECONDS)

    return None, last_error


def save_json(data: Dict[str, Any], path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def normalize_csv_header(fieldnames: Optional[List[str]]) -> List[str]:
    """
    Прибирає BOM з першої колонки CSV, якщо файл збережено як UTF-8 BOM.
    """
    if not fieldnames:
        return []

    cleaned = []
    for name in fieldnames:
        if name is None:
            cleaned.append("")
        else:
            cleaned.append(name.replace("\ufeff", "").strip())
    return cleaned


def has_valid_tender_ids(rows: List[Dict[str, str]]) -> bool:
    """
    Перевіряє, чи CSV реально прочитав колонку tenderID, а не зсунуті поля.
    """
    return any(PUBLIC_TENDER_ID_RE.fullmatch(str(row.get("tenderID", ""))) for row in rows)


def read_input_rows_by_regex(input_path: Path) -> List[Dict[str, str]]:
    """
    Резервний парсер для пошкоджених CSV з неекранованими переносами рядків.
    Для завантаження документів потрібні тільки id та tenderID.
    """
    text = input_path.read_bytes().decode("latin-1")
    pairs = re.findall(
        rf"({INTERNAL_ID_RE.pattern})(?:(?!{INTERNAL_ID_RE.pattern}).)*?({PUBLIC_TENDER_ID_RE.pattern})",
        text,
        flags=re.S,
    )

    rows = []
    seen = set()
    for internal_id, public_id in pairs:
        key = (internal_id, public_id)
        if key in seen:
            continue
        seen.add(key)
        rows.append({"id": internal_id, "tenderID": public_id})

    if not rows:
        raise ValueError("Не вдалося витягти пари id/tenderID з пошкодженого CSV")

    print(
        "[!] CSV має некоректну структуру для стандартного парсера; "
        f"витягнуто пар id/tenderID регулярним виразом: {len(rows)}"
    )
    return rows


def get_nested_documents(tender_data: Dict[str, Any]) -> List[Tuple[str, Dict[str, Any]]]:
    """
    Збирає документи з верхнього рівня та ключових вкладених блоків Prozorro.
    Повертає список: (section_path, document_dict).
    """
    found: List[Tuple[str, Dict[str, Any]]] = []

    def add_documents(section: str, docs: Any) -> None:
        if isinstance(docs, list):
            for doc in docs:
                if isinstance(doc, dict):
                    found.append((section, doc))

    add_documents("documents", tender_data.get("documents"))

    for i, complaint in enumerate(tender_data.get("complaints", []) or []):
        if isinstance(complaint, dict):
            add_documents(f"complaints/{i}/documents", complaint.get("documents"))

    for i, lot in enumerate(tender_data.get("lots", []) or []):
        if isinstance(lot, dict):
            add_documents(f"lots/{i}/documents", lot.get("documents"))

    for i, bid in enumerate(tender_data.get("bids", []) or []):
        if isinstance(bid, dict):
            add_documents(f"bids/{i}/documents", bid.get("documents"))

    for i, award in enumerate(tender_data.get("awards", []) or []):
        if not isinstance(award, dict):
            continue

        add_documents(f"awards/{i}/documents", award.get("documents"))

        for j, complaint in enumerate(award.get("complaints", []) or []):
            if isinstance(complaint, dict):
                add_documents(
                    f"awards/{i}/complaints/{j}/documents",
                    complaint.get("documents"),
                )

    for i, qualification in enumerate(tender_data.get("qualifications", []) or []):
        if not isinstance(qualification, dict):
            continue

        add_documents(f"qualifications/{i}/documents", qualification.get("documents"))

        for j, milestone in enumerate(qualification.get("milestones", []) or []):
            if isinstance(milestone, dict):
                add_documents(
                    f"qualifications/{i}/milestones/{j}/documents",
                    milestone.get("documents"),
                )

        for j, complaint in enumerate(qualification.get("complaints", []) or []):
            if isinstance(complaint, dict):
                add_documents(
                    f"qualifications/{i}/complaints/{j}/documents",
                    complaint.get("documents"),
                )

    for i, cancellation in enumerate(tender_data.get("cancellations", []) or []):
        if not isinstance(cancellation, dict):
            continue

        add_documents(f"cancellations/{i}/documents", cancellation.get("documents"))

        for j, complaint in enumerate(cancellation.get("complaints", []) or []):
            if isinstance(complaint, dict):
                add_documents(
                    f"cancellations/{i}/complaints/{j}/documents",
                    complaint.get("documents"),
                )

    for i, contract in enumerate(tender_data.get("contracts", []) or []):
        if not isinstance(contract, dict):
            continue

        add_documents(f"contracts/{i}/documents", contract.get("documents"))

        for j, change in enumerate(contract.get("changes", []) or []):
            if isinstance(change, dict):
                add_documents(
                    f"contracts/{i}/changes/{j}/documents",
                    change.get("documents"),
                )

    return found


def section_to_folder_name(section: str) -> str:
    """
    Перетворює section_path на безпечну назву папки.
    """
    return safe_filename(section.replace("/", "__"), max_len=160)


def guess_extension_from_response(response: requests.Response, original_name: str) -> str:
    """
    Якщо в title немає розширення, пробує визначити його з Content-Type.
    """
    if Path(original_name).suffix:
        return ""

    content_type = response.headers.get("Content-Type", "").split(";")[0].strip().lower()

    mapping = {
        "application/pdf": ".pdf",
        "application/zip": ".zip",
        "application/x-zip-compressed": ".zip",
        "application/msword": ".doc",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document": ".docx",
        "application/vnd.ms-excel": ".xls",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": ".xlsx",
        "application/vnd.ms-powerpoint": ".ppt",
        "application/vnd.openxmlformats-officedocument.presentationml.presentation": ".pptx",
        "text/plain": ".txt",
        "text/csv": ".csv",
        "image/jpeg": ".jpg",
        "image/png": ".png",
    }

    return mapping.get(content_type, "")


def download_document(
    *,
    doc: Dict[str, Any],
    section: str,
    tender_folder: Path,
    public_id: str,
    internal_id: str,
) -> Dict[str, Any]:
    """
    Завантажує один документ.
    Повертає запис для лог-файлу.
    """
    doc_id = str(doc.get("id") or "")
    title = str(doc.get("title") or "document")
    download_url = doc.get("url")

    log_row = {
        "tenderID": public_id,
        "internal_id": internal_id,
        "section": section,
        "document_id": doc_id,
        "title": title,
        "url": download_url or "",
        "status": "",
        "local_path": "",
        "error": "",
    }

    if not download_url:
        log_row["status"] = "skipped"
        log_row["error"] = "missing document url"
        return log_row

    response, error = request_with_retries(
        str(download_url),
        timeout=REQUEST_TIMEOUT_FILE,
        expect_json=False,
    )

    if error or response is None:
        log_row["status"] = "failed"
        log_row["error"] = error or "unknown request error"
        return log_row

    section_folder = tender_folder / "files" / section_to_folder_name(section)
    section_folder.mkdir(parents=True, exist_ok=True)

    safe_title = safe_filename(title)
    prefix = safe_filename(doc_id, max_len=60) if doc_id else "no_doc_id"
    extension = guess_extension_from_response(response, safe_title)
    final_name = f"{prefix}_{safe_title}{extension}"

    file_path = ensure_unique_path(section_folder / final_name)

    try:
        with file_path.open("wb") as f:
            f.write(response.content)

        log_row["status"] = "downloaded"
        log_row["local_path"] = str(file_path)

    except Exception as e:
        log_row["status"] = "failed"
        log_row["error"] = f"file write error: {e}"

    return log_row


# =============================================================================
# ОСНОВНА ЛОГІКА
# =============================================================================

def fetch_tender_full_json(internal_id: str) -> Tuple[Optional[Dict[str, Any]], Optional[str]]:
    """
    Завантажує повний JSON тендера за внутрішнім Prozorro ID.
    """
    url = f"{PROZORRO_API_BASE}/tenders/{internal_id}"

    response, error = request_with_retries(
        url,
        timeout=REQUEST_TIMEOUT_TENDER,
        expect_json=True,
    )

    if error or response is None:
        return None, error or "unknown tender request error"

    try:
        payload = response.json()
    except Exception as e:
        return None, f"json parse error: {e}"

    tender_data = payload.get("data")
    if not isinstance(tender_data, dict):
        return None, "response has no data object"

    return tender_data, None


def process_tender(internal_id: str, public_id: str) -> List[Dict[str, Any]]:
    """
    Обробляє один тендер:
    - отримує повний JSON;
    - зберігає tender_full.json;
    - зберігає tender_summary.json;
    - завантажує всі знайдені документи з основних вкладених блоків;
    - повертає лог-рядки.
    """
    internal_id = str(internal_id).strip()
    public_id = str(public_id).strip()

    tender_folder = Path(DOWNLOAD_DIR) / safe_filename(public_id, max_len=120)
    tender_folder.mkdir(parents=True, exist_ok=True)

    print(f"\n=== {public_id} ===")
    print(f"internal id: {internal_id}")

    tender_data, error = fetch_tender_full_json(internal_id)

    if error or tender_data is None:
        print(f"[!] Не вдалося отримати JSON тендера: {error}")

        return [{
            "tenderID": public_id,
            "internal_id": internal_id,
            "section": "tender_json",
            "document_id": "",
            "title": "tender_full.json",
            "url": f"{PROZORRO_API_BASE}/tenders/{internal_id}",
            "status": "failed",
            "local_path": "",
            "error": error or "unknown error",
        }]

    full_json_path = tender_folder / "tender_full.json"
    save_json(tender_data, full_json_path)

    summary = {
        "id": tender_data.get("id"),
        "tenderID": tender_data.get("tenderID"),
        "status": tender_data.get("status"),
        "procurementMethod": tender_data.get("procurementMethod"),
        "procurementMethodType": tender_data.get("procurementMethodType"),
        "mainProcurementCategory": tender_data.get("mainProcurementCategory"),
        "title": tender_data.get("title"),
        "value": tender_data.get("value"),
        "procuringEntity": tender_data.get("procuringEntity"),
        "number_of_lots": len(tender_data.get("lots", []) or []),
        "number_of_bids": len(tender_data.get("bids", []) or []),
        "number_of_awards": len(tender_data.get("awards", []) or []),
        "number_of_contracts": len(tender_data.get("contracts", []) or []),
        "number_of_complaints": len(tender_data.get("complaints", []) or []),
    }
    save_json(summary, tender_folder / "tender_summary.json")

    print("[+] JSON збережено")
    print(f"    status: {summary.get('status')}")
    print(f"    methodType: {summary.get('procurementMethodType')}")
    print(f"    category: {summary.get('mainProcurementCategory')}")
    print(f"    bids: {summary.get('number_of_bids')}")
    print(f"    awards: {summary.get('number_of_awards')}")
    print(f"    contracts: {summary.get('number_of_contracts')}")

    docs = get_nested_documents(tender_data)
    print(f"[+] Знайдено документів у всіх перевірених блоках: {len(docs)}")

    log_rows: List[Dict[str, Any]] = []

    if not docs:
        log_rows.append({
            "tenderID": public_id,
            "internal_id": internal_id,
            "section": "documents",
            "document_id": "",
            "title": "",
            "url": "",
            "status": "no_documents",
            "local_path": "",
            "error": "",
        })
        return log_rows

    for section, doc in docs:
        row = download_document(
            doc=doc,
            section=section,
            tender_folder=tender_folder,
            public_id=public_id,
            internal_id=internal_id,
        )
        log_rows.append(row)

        if row["status"] == "downloaded":
            print(f"    [+] {section}: {row['title']}")
        else:
            print(f"    [!] {section}: {row['title']} — {row['error']}")

    return log_rows


def read_input_rows(input_csv: Path) -> List[Dict[str, str]]:
    """
    Читає CSV і повертає рядки з нормалізованими ключами.
    """
    input_path = Path(input_csv)

    if not input_path.exists():
        raise FileNotFoundError(f"Файл не знайдено: {input_path}")

    decode_errors = []

    for encoding in ("utf-8-sig", "cp1251"):
        try:
            with input_path.open(mode="r", encoding=encoding, newline="") as f:
                reader = csv.DictReader(f)

                if not reader.fieldnames:
                    raise ValueError("CSV не має заголовків")

                original_fieldnames = reader.fieldnames
                cleaned_fieldnames = normalize_csv_header(original_fieldnames)
                name_map = dict(zip(original_fieldnames, cleaned_fieldnames))

                rows: List[Dict[str, str]] = []
                for raw_row in reader:
                    cleaned_row = {
                        name_map.get(k, k): (v.strip() if isinstance(v, str) else v)
                        for k, v in raw_row.items()
                    }
                    rows.append(cleaned_row)

            break
        except UnicodeDecodeError as e:
            decode_errors.append(f"{encoding}: {e}")
    else:
        raise UnicodeDecodeError(
            "csv",
            b"",
            0,
            1,
            "Не вдалося прочитати CSV у кодуваннях utf-8-sig або cp1251. "
            + " | ".join(decode_errors),
        )

    required_fields = ["id", "tenderID"]
    missing = [field for field in required_fields if field not in cleaned_fieldnames]

    if missing:
        raise ValueError(
            f"CSV файл повинен містити колонки {required_fields}. "
            f"Відсутні: {missing}. "
            f"Доступні колонки: {cleaned_fieldnames}"
        )

    if not has_valid_tender_ids(rows):
        return read_input_rows_by_regex(input_path)

    return rows


def write_download_log(log_rows: List[Dict[str, Any]], path: Path) -> None:
    """
    Записує загальний лог завантаження.
    """
    path.parent.mkdir(parents=True, exist_ok=True)

    fieldnames = [
        "tenderID",
        "internal_id",
        "section",
        "document_id",
        "title",
        "url",
        "status",
        "local_path",
        "error",
    ]

    with path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()

        for row in log_rows:
            writer.writerow({field: row.get(field, "") for field in fieldnames})


def main() -> None:
    print("=== Запуск завантаження повних пакетів тендерів Prozorro ===")

    Path(DOWNLOAD_DIR).mkdir(parents=True, exist_ok=True)

    try:
        rows = read_input_rows(INPUT_CSV)
    except Exception as e:
        print(f"[Критична помилка] {e}")
        return

    print(f"[+] Прочитано рядків з CSV: {len(rows)}")

    all_log_rows: List[Dict[str, Any]] = []

    for index, row in enumerate(rows, start=1):
        internal_id = row.get("id", "")
        public_id = row.get("tenderID", "")

        if not internal_id or not public_id:
            print(f"[!] Рядок {index}: пропущено, бо немає id або tenderID")
            all_log_rows.append({
                "tenderID": public_id,
                "internal_id": internal_id,
                "section": "input",
                "document_id": "",
                "title": "",
                "url": "",
                "status": "skipped",
                "local_path": "",
                "error": "missing id or tenderID",
            })
            continue

        tender_log_rows = process_tender(internal_id, public_id)
        all_log_rows.extend(tender_log_rows)

        if any(
            str(log_row.get("error", "")).startswith(FATAL_NETWORK_PREFIX)
            for log_row in tender_log_rows
        ):
            print("[Критична помилка] Немає DNS/мережевого доступу до Prozorro API. Пакет зупинено.")
            break

    log_path = Path(DOWNLOAD_DIR) / "download_log.csv"
    write_download_log(all_log_rows, log_path)

    print("\n=== Обробку завершено ===")
    print(f"[+] Лог збережено: {log_path}")


if __name__ == "__main__":
    main()
