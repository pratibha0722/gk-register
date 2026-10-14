#!/usr/bin/env python3
"""
Shared extraction engine for handwritten/scanned GK register images/PDFs.

Visible source columns (preserved in order):
A  S.No
B  Date
C  G.K. No
D  UHID No
E  Name of Patient
F  Age/Sex
G  Diagnosis
H  Dose
I  Perc
J  Volume
K  GURA
L  CONS
M  PHY
N  R.O
O  RTT
P  Technique (derived from diagnosis marker: (M)=Masked, (V)=Vantage)
Q:R blank
S  File Name

Key behavior:
- Gemini multimodal extraction from images or one PDF page at a time.
- Two Gemini API keys with quota/high-demand failover.
- Deterministic post-processing for GK repeat marker and diagnosis technique marker.
- Local Excel UPSERT, not blind append: S.No is the primary identity; Date is
  secondary/guard information. Re-uploading the same row updates it instead of
  creating another row.
- If the same S.No is seen with a different nonblank date, the existing date is
  preserved to prevent a hallucinated date from creating/altering a duplicate.
- Google Apps Script writer uses dataset='gk' and expects the companion combined
  receiver google_sheet_receiver_ot_gk_v10.gs.
"""

from __future__ import annotations

import hashlib
import io
import re
import time
from dataclasses import dataclass
from datetime import datetime, time as dt_time, timedelta
from pathlib import Path
from typing import Callable, Iterable, List, Optional
from zoneinfo import ZoneInfo

import pymupdf
import requests
from google import genai
from google.genai import errors as genai_errors
from google.genai import types
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from PIL import Image, ImageOps
from pydantic import BaseModel, Field

try:
    import pillow_heif
    pillow_heif.register_heif_opener()
except Exception:
    pillow_heif = None


# -----------------------------------------------------------------------------
# Schema
# -----------------------------------------------------------------------------

SOURCE_FIELDS = [
    "sno",
    "date",
    "gk_no",
    "uhid_no",
    "patient_name",
    "age_sex",
    "diagnosis",
    "dose",
    "perc",
    "volume",
    "gura",
    "cons",
    "phy",
    "ro",
    "rtt",
]

EXCEL_HEADERS = [
    "S.No",
    "Date",
    "G.K. No",
    "UHID No",
    "Name of Patient",
    "Age/Sex",
    "Diagnosis",
    "Dose",
    "Perc",
    "Volume",
    "GURA",
    "CONS",
    "PHY",
    "R.O",
    "RTT",
    "Technique",
]

TECHNIQUE_COL = 16  # P
FILENAME_COL = 19   # S; Q:R intentionally blank
FILENAME_HEADER = "File Name"


class GKRecord(BaseModel):
    source_row: int = Field(
        default=0,
        description="Physical patient/data row number on the page, top-to-bottom starting at 1.",
    )
    sno: str = ""
    date: str = ""
    gk_no: str = ""
    uhid_no: str = ""
    patient_name: str = ""
    age_sex: str = ""
    diagnosis: str = ""
    dose: str = ""
    perc: str = ""
    volume: str = ""
    gura: str = ""
    cons: str = ""
    phy: str = ""
    ro: str = ""
    rtt: str = ""
    technique: str = ""  # derived locally; model may leave blank


class GKPageData(BaseModel):
    records: List[GKRecord] = Field(default_factory=list)


PROMPT = r"""
You are transcribing ONE PAGE/IMAGE of a handwritten or scanned Gamma Knife
register. Read the physical table visually. Accuracy and row alignment are more
important than guessing.

The logical columns are normally, left-to-right:
1. S.No
2. Date
3. G.K. No
4. UHID No
5. Name of Patient
6. Age/Sex
7. Diagnosis
8. Dose
9. Perc
10. Volume
11. GURA
12. CONS
13. PHY
14. R.O
15. RTT

The right-side abbreviations may be handwritten and narrow. Keep their values in
the correct row and column. Do not move a value to a neighboring row.

LAYOUT / ORIENTATION:
- Inspect THIS image/page independently. Do not assume fixed coordinates.
- The image can be photographed, skewed, upside-down, or rotated 90/180/270°.
  Interpret it in the correct visual orientation before reading it.
- The table can span an open book across the center fold.
- Ignore camera timestamp overlays, page margins, titles/month headings, and
  printed ruling unless they are actual row values.

ONE PHYSICAL PATIENT ROW = ONE RECORD:
- Return every real data row exactly once, top-to-bottom.
- source_row is the physical row number on this page starting at 1.
- Never create two records for the same physical row.
- If text wraps to a second handwritten line inside the same ruled row, combine
  it into the same record.
- Omit a row only if the whole row is clearly cancelled/crossed out.

S.No AND DATE ARE CRITICAL:
- Copy S.No only from that row's S.No cell.
- Copy Date only from that row's Date cell.
- NEVER invent/carry a date from the row above/below, month heading, page title,
  sequence, S.No, GK number, or any other context.
- If a date is unreadable/blank, return "" rather than guessing.
- Preserve visible punctuation/style such as 08-09-26 or 8.09.26.

G.K. REPEAT RULE:
- Transcribe the G.K. cell literally enough to preserve a repeat marker.
- If the number has an R/(R)/r repeat marker, keep both number and marker in
  gk_no (examples: "7282 (R)", "R 7282"). Local code will normalize this to
  "Repeat 7282".
- Never invent a repeat marker.

DIAGNOSIS MARKER RULE:
- Preserve visible (M) or (V) inside diagnosis exactly when present.
- (M) means Masked and (V) means Vantage; local code will derive Technique.
- Do not invent M/V if it is not visibly written.

TRANSCRIPTION:
- Copy handwriting literally; do not medically 'correct' names/diagnoses.
- Do not infer missing values from nearby rows.
- Keep numbers and units such as Gy, %, cc, mm exactly when legible.
- For unreadable cells, return "" rather than hallucinating.
- Do not include month headings such as SEPTEMBER 2026 as patient rows.
"""


# -----------------------------------------------------------------------------
# Deterministic normalization / duplicate prevention
# -----------------------------------------------------------------------------


def clean_text(value: object) -> str:
    text = "" if value is None else str(value)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def normalize_sno(value: str) -> str:
    s = clean_text(value)
    s = re.sub(r"^[#\s]+", "", s)
    s = re.sub(r"[.\s]+$", "", s)
    return s


def normalize_gk_repeat(value: str) -> str:
    s = clean_text(value)
    if not s:
        return ""
    if re.match(r"(?i)^repeat\b", s):
        # Normalize spacing but preserve ID text after Repeat.
        rest = re.sub(r"(?i)^repeat\s*[:\-]?\s*", "", s).strip()
        return f"Repeat {rest}".strip()

    # Common forms: 7282 (R), 7282 R, R 7282, (R) 7282.
    has_r = bool(
        re.search(r"(?i)\(\s*r\s*\)", s)
        or re.search(r"(?i)(?:^|\s)r(?:\s|$)", s)
        or re.match(r"(?i)^r\s*[-:]?\s*\d", s)
    )
    if not has_r:
        return s

    without_marker = re.sub(r"(?i)\(\s*r\s*\)", " ", s)
    without_marker = re.sub(r"(?i)(?:^|\s)r(?:\s|$)", " ", without_marker)
    without_marker = re.sub(r"\s+", " ", without_marker).strip(" :-")
    return f"Repeat {without_marker}".strip()


def diagnosis_and_technique(value: str) -> tuple[str, str]:
    s = clean_text(value)
    technique = ""
    if re.search(r"(?i)\(\s*m\s*\)", s):
        technique = "Masked"
        s = re.sub(r"(?i)\(\s*m\s*\)", " ", s)
    if re.search(r"(?i)\(\s*v\s*\)", s):
        technique = "Vantage" if not technique else technique + " / Vantage"
        s = re.sub(r"(?i)\(\s*v\s*\)", " ", s)
    s = re.sub(r"\s+", " ", s).strip(" -;,|")
    return s, technique


def normalize_record(record: GKRecord) -> GKRecord:
    data = record.model_dump()
    for k in list(data):
        if k != "source_row":
            data[k] = clean_text(data[k])
    data["sno"] = normalize_sno(data["sno"])
    data["gk_no"] = normalize_gk_repeat(data["gk_no"])
    diagnosis, technique = diagnosis_and_technique(data["diagnosis"])
    data["diagnosis"] = diagnosis
    data["technique"] = technique
    return GKRecord(**data)


def merge_records(a: GKRecord, b: GKRecord) -> GKRecord:
    """Merge duplicate interpretations without letting a new date overwrite one."""
    out = a.model_dump()
    incoming = b.model_dump()
    for field in SOURCE_FIELDS + ["technique"]:
        av = clean_text(out.get(field, ""))
        bv = clean_text(incoming.get(field, ""))
        if field == "date" and av and bv and av != bv:
            # Keep first defensible row date; a disagreement must not make a duplicate.
            continue
        if not av and bv:
            out[field] = bv
    if not out.get("source_row") and incoming.get("source_row"):
        out["source_row"] = incoming["source_row"]
    return GKRecord(**out)


def dedupe_records(records: Iterable[GKRecord]) -> List[GKRecord]:
    output: List[GKRecord] = []
    by_source: dict[int, int] = {}
    by_sno: dict[str, int] = {}

    for raw in records:
        r = normalize_record(raw)
        if not any(clean_text(getattr(r, f)) for f in SOURCE_FIELDS):
            continue

        idx: Optional[int] = None
        if r.source_row > 0 and r.source_row in by_source:
            idx = by_source[r.source_row]
        sno_key = normalize_sno(r.sno).casefold()
        if idx is None and sno_key and sno_key in by_sno:
            idx = by_sno[sno_key]

        if idx is None:
            idx = len(output)
            output.append(r)
        else:
            output[idx] = merge_records(output[idx], r)

        if r.source_row > 0:
            by_source[r.source_row] = idx
        if sno_key:
            by_sno[sno_key] = idx

    output.sort(key=lambda r: r.source_row if r.source_row > 0 else 10**9)
    return output


def record_to_row(r: GKRecord) -> List[str]:
    return [
        r.sno, r.date, r.gk_no, r.uhid_no, r.patient_name, r.age_sex,
        r.diagnosis, r.dose, r.perc, r.volume, r.gura, r.cons, r.phy,
        r.ro, r.rtt, r.technique,
    ]


# -----------------------------------------------------------------------------
# Gemini quota/failover
# -----------------------------------------------------------------------------

DAILY_RESET_TIMEZONE = "America/Los_Angeles"
DAILY_RESET_GRACE_MINUTES = 5
GENERIC_429_SHORT_WAITS = (30, 60, 120)
SERVER_HIGH_DEMAND_WAIT_SECONDS = 10 * 60


@dataclass
class UsageTotals:
    input_tokens: int = 0
    output_tokens: int = 0
    thinking_tokens: int = 0
    total_tokens: int = 0

    def add_response(self, response: object) -> None:
        usage = getattr(response, "usage_metadata", None)
        if usage is None:
            return
        self.input_tokens += int(getattr(usage, "prompt_token_count", 0) or 0)
        self.output_tokens += int(getattr(usage, "candidates_token_count", 0) or 0)
        self.thinking_tokens += int(getattr(usage, "thoughts_token_count", 0) or 0)
        self.total_tokens += int(getattr(usage, "total_token_count", 0) or 0)


class QuotaUnavailable(RuntimeError):
    pass


class GeminiKeyPool:
    def __init__(self, keys: List[str]):
        cleaned: List[str] = []
        for raw in keys:
            key = clean_text(raw)
            if not key or "PASTE_YOUR" in key:
                continue
            if key not in cleaned:
                cleaned.append(key)
        if not cleaned:
            raise ValueError("No usable Gemini API key was provided")
        self._clients = [genai.Client(api_key=k) for k in cleaned]
        self._current = 0
        self._daily_blocked: set[int] = set()

    @property
    def count(self) -> int:
        return len(self._clients)

    @property
    def current_index(self) -> int:
        return self._current

    @property
    def current_client(self):
        return self._clients[self._current]

    @property
    def current_label(self) -> str:
        return f"Gemini API {self._current + 1}"

    def mark_daily_blocked(self) -> None:
        self._daily_blocked.add(self._current)

    def switch_to_next_available(self, extra_blocked: Optional[set[int]] = None) -> bool:
        blocked = set(self._daily_blocked)
        if extra_blocked:
            blocked.update(extra_blocked)
        for offset in range(1, self.count + 1):
            candidate = (self._current + offset) % self.count
            if candidate not in blocked:
                self._current = candidate
                return True
        return False

    def reset_daily(self) -> None:
        self._daily_blocked.clear()
        self._current = 0


def api_error_code(exc: Exception) -> Optional[int]:
    code = getattr(exc, "code", None)
    try:
        return int(code) if code is not None else None
    except Exception:
        return None


def api_error_text(exc: Exception) -> str:
    msg = getattr(exc, "message", None)
    return str(msg if msg is not None else exc)


def is_daily_quota_error(exc: Exception) -> bool:
    if api_error_code(exc) != 429:
        return False
    text = api_error_text(exc).lower().replace("-", "_")
    markers = (
        "quota_exceeded", "daily quota", "per day", "requests per day",
        "request per day", "tokens per day", "rpd", "tpd",
        "requestsperday", "tokensperday", "free_tier_requests",
        "free tier requests", "generate_content_free_tier_requests",
    )
    return any(m in text for m in markers)


def is_429(exc: Exception) -> bool:
    return api_error_code(exc) == 429 or "resource_exhausted" in api_error_text(exc).lower()


def is_retryable_server_error(exc: Exception) -> bool:
    return api_error_code(exc) in {500, 502, 503, 504}


def seconds_until_next_pacific_reset() -> tuple[int, datetime]:
    tz = ZoneInfo(DAILY_RESET_TIMEZONE)
    now = datetime.now(tz)
    tomorrow = now.date() + timedelta(days=1)
    reset = datetime.combine(tomorrow, dt_time(0, DAILY_RESET_GRACE_MINUTES), tzinfo=tz)
    return max(1, int((reset - now).total_seconds())), reset


def media_resolution_enum(name: str):
    return {
        "low": types.MediaResolution.MEDIA_RESOLUTION_LOW,
        "medium": types.MediaResolution.MEDIA_RESOLUTION_MEDIUM,
        "high": types.MediaResolution.MEDIA_RESOLUTION_HIGH,
    }[name]


def _sleep(seconds: int, log: Callable[[str], None]) -> None:
    end = time.time() + max(0, seconds)
    while True:
        remaining = int(end - time.time())
        if remaining <= 0:
            return
        if remaining > 3600:
            log(f"waiting: about {remaining//3600}h {(remaining%3600)//60}m remaining")
            chunk = min(1800, remaining)
        elif remaining > 60:
            log(f"waiting: about {remaining//60}m remaining")
            chunk = min(60, remaining)
        else:
            log(f"waiting: about {remaining}s remaining")
            chunk = min(10, remaining)
        time.sleep(chunk)


def generate_with_failover(
    key_pool: GeminiKeyPool,
    model: str,
    contents: list,
    config: types.GenerateContentConfig,
    retries: int = 3,
    wait_on_quota: bool = True,
    log: Callable[[str], None] = print,
    status_callback: Optional[Callable[[str, str], None]] = None,
):
    generic_429_count = 0
    server_attempt = 0
    temporary_server_blocked: set[int] = set()

    def status(state: str, message: str) -> None:
        if status_callback:
            try:
                status_callback(state, message)
            except Exception:
                pass

    while True:
        active_index = key_pool.current_index
        active_label = key_pool.current_label
        try:
            response = key_pool.current_client.models.generate_content(
                model=model,
                contents=contents,
                config=config,
            )
            temporary_server_blocked.clear()
            return response
        except KeyboardInterrupt:
            raise
        except genai_errors.APIError as exc:
            text = api_error_text(exc)

            if is_daily_quota_error(exc):
                log(f"{active_label}: daily/free-tier quota unavailable")
                key_pool.mark_daily_blocked()
                temporary_server_blocked.discard(active_index)
                if key_pool.switch_to_next_available(extra_blocked=temporary_server_blocked):
                    log(f"switching automatically to {key_pool.current_label}; retrying same page/image")
                    status("Processing", f"{active_label} quota consumed; continuing with {key_pool.current_label}.")
                    generic_429_count = 0
                    server_attempt = 0
                    continue

                if not wait_on_quota:
                    raise QuotaUnavailable(
                        "All configured Gemini API projects have exhausted their current daily quota."
                    ) from exc

                seconds, reset_pt = seconds_until_next_pacific_reset()
                local_reset = reset_pt.astimezone()
                msg = "Gemini quota consumed. Job is paused and will restart automatically after reset."
                log(msg)
                log(f"next automatic retry: {local_reset.strftime('%Y-%m-%d %H:%M %Z')}")
                status("Waiting for Gemini quota", msg)
                _sleep(seconds, log)
                key_pool.reset_daily()
                temporary_server_blocked.clear()
                status("Processing", "Gemini quota reset window reached; retrying automatically.")
                generic_429_count = 0
                server_attempt = 0
                continue

            if is_429(exc):
                generic_429_count += 1
                if generic_429_count <= len(GENERIC_429_SHORT_WAITS):
                    wait = GENERIC_429_SHORT_WAITS[generic_429_count - 1]
                    log(f"{active_label} 429/rate limit ({generic_429_count}/3): waiting {wait}s")
                    status("Gemini rate limited", f"{active_label} temporarily rate limited; retrying.")
                    _sleep(wait, log)
                    continue
                key_pool.mark_daily_blocked()
                temporary_server_blocked.discard(active_index)
                if key_pool.switch_to_next_available(extra_blocked=temporary_server_blocked):
                    log(f"429 persisted; switching to {key_pool.current_label}")
                    generic_429_count = 0
                    server_attempt = 0
                    continue
                if not wait_on_quota:
                    raise QuotaUnavailable("All Gemini APIs are currently rate/quota limited.") from exc
                seconds, _ = seconds_until_next_pacific_reset()
                status("Waiting for Gemini quota", "All Gemini APIs are rate/quota limited; waiting for reset.")
                _sleep(seconds, log)
                key_pool.reset_daily()
                temporary_server_blocked.clear()
                generic_429_count = 0
                server_attempt = 0
                continue

            if is_retryable_server_error(exc):
                server_attempt += 1
                if server_attempt <= max(1, retries):
                    wait = min(2 ** server_attempt, 30)
                    log(f"{active_label} server/high-demand error {api_error_code(exc)} ({server_attempt}/{retries}); retrying in {wait}s")
                    status("Gemini high demand", f"{active_label} temporarily busy; retrying.")
                    _sleep(wait, log)
                    continue
                temporary_server_blocked.add(active_index)
                if key_pool.switch_to_next_available(extra_blocked=temporary_server_blocked):
                    log(f"high demand persisted; switching to {key_pool.current_label}")
                    server_attempt = 0
                    generic_429_count = 0
                    continue
                if not wait_on_quota:
                    raise RuntimeError("All Gemini APIs are temporarily experiencing high demand.") from exc
                msg = "All Gemini APIs are temporarily experiencing high demand; waiting 10 minutes."
                status("Waiting for Gemini service", msg)
                log(msg)
                _sleep(SERVER_HIGH_DEMAND_WAIT_SECONDS, log)
                temporary_server_blocked.clear()
                server_attempt = 0
                generic_429_count = 0
                continue
            raise


def parse_page_response(response: object) -> GKPageData:
    parsed = getattr(response, "parsed", None)
    if isinstance(parsed, GKPageData):
        return parsed
    if parsed is not None:
        try:
            return GKPageData.model_validate(parsed)
        except Exception:
            pass
    text = getattr(response, "text", None)
    if not text:
        raise RuntimeError("Gemini returned an empty structured response")
    return GKPageData.model_validate_json(text)


def extract_page(
    key_pool: GeminiKeyPool,
    media_bytes: bytes,
    mime_type: str,
    model: str = "gemini-3.5-flash",
    media_resolution: str = "high",
    thinking_level: str = "low",
    retries: int = 3,
    usage: Optional[UsageTotals] = None,
    wait_on_quota: bool = True,
    log: Callable[[str], None] = print,
    status_callback: Optional[Callable[[str, str], None]] = None,
) -> List[GKRecord]:
    part = types.Part.from_bytes(data=media_bytes, mime_type=mime_type)
    config = types.GenerateContentConfig(
        response_mime_type="application/json",
        response_schema=GKPageData,
        media_resolution=media_resolution_enum(media_resolution),
        thinking_config=types.ThinkingConfig(thinking_level=thinking_level),
    )

    last_exc: Optional[Exception] = None
    for parse_attempt in range(1, max(1, retries) + 1):
        response = generate_with_failover(
            key_pool=key_pool,
            model=model,
            contents=[part, PROMPT],
            config=config,
            retries=retries,
            wait_on_quota=wait_on_quota,
            log=log,
            status_callback=status_callback,
        )
        if usage:
            usage.add_response(response)
        try:
            return dedupe_records(parse_page_response(response).records)
        except Exception as exc:
            last_exc = exc
            if parse_attempt < retries:
                log(f"structured response parse failed ({parse_attempt}/{retries}); retrying same page/image")
                continue
            break
    raise RuntimeError(f"Gemini structured extraction failed: {last_exc}")


# -----------------------------------------------------------------------------
# Media preparation
# -----------------------------------------------------------------------------

IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".webp", ".heic", ".heif", ".tif", ".tiff"}


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def image_to_jpeg_bytes(path: Path) -> bytes:
    with Image.open(path) as im:
        im = ImageOps.exif_transpose(im)
        if im.mode not in ("RGB", "L"):
            im = im.convert("RGB")
        elif im.mode == "L":
            im = im.convert("RGB")
        buf = io.BytesIO()
        im.save(buf, format="JPEG", quality=95, optimize=True)
        return buf.getvalue()


def uploaded_image_to_jpeg_bytes(data: bytes) -> bytes:
    with Image.open(io.BytesIO(data)) as im:
        im = ImageOps.exif_transpose(im)
        if im.mode != "RGB":
            im = im.convert("RGB")
        buf = io.BytesIO()
        im.save(buf, format="JPEG", quality=95, optimize=True)
        return buf.getvalue()


def one_page_pdf_bytes(doc: pymupdf.Document, page_index: int) -> bytes:
    single = pymupdf.open()
    try:
        single.insert_pdf(doc, from_page=page_index, to_page=page_index)
        return single.tobytes(garbage=4, deflate=True)
    finally:
        single.close()


def iter_path_pages(path: Path):
    ext = path.suffix.lower()
    if ext == ".pdf":
        doc = pymupdf.open(str(path))
        try:
            total = len(doc)
            for i in range(total):
                payload = one_page_pdf_bytes(doc, i)
                yield i + 1, total, payload, "application/pdf", sha256_bytes(payload)
        finally:
            doc.close()
    elif ext in IMAGE_EXTS:
        payload = image_to_jpeg_bytes(path)
        yield 1, 1, payload, "image/jpeg", sha256_bytes(payload)
    else:
        raise ValueError(f"Unsupported file type: {path.suffix}")


def iter_uploaded_pages(filename: str, data: bytes):
    ext = Path(filename).suffix.lower()
    if ext == ".pdf":
        doc = pymupdf.open(stream=data, filetype="pdf")
        try:
            total = len(doc)
            for i in range(total):
                payload = one_page_pdf_bytes(doc, i)
                yield i + 1, total, payload, "application/pdf", sha256_bytes(payload)
        finally:
            doc.close()
    elif ext in IMAGE_EXTS:
        payload = uploaded_image_to_jpeg_bytes(data)
        yield 1, 1, payload, "image/jpeg", sha256_bytes(payload)
    else:
        raise ValueError(f"Unsupported file type: {ext}")


# -----------------------------------------------------------------------------
# Excel UPSERT
# -----------------------------------------------------------------------------


def _date_norm(value: object) -> str:
    return re.sub(r"\s+", "", clean_text(value)).casefold()


def _fallback_identity(row: List[str]) -> str:
    # Date + GK + UHID + patient name is used only when S.No is missing.
    pieces = [row[1], row[2], row[3], row[4]]
    return "|".join(clean_text(x).casefold() for x in pieces)


def ensure_excel_layout(ws) -> None:
    for col, header in enumerate(EXCEL_HEADERS, start=1):
        if ws.cell(1, col).value in (None, ""):
            ws.cell(1, col, header)
    if ws.cell(1, FILENAME_COL).value in (None, ""):
        ws.cell(1, FILENAME_COL, FILENAME_HEADER)

    fill = PatternFill("solid", fgColor="1F4E78")
    font = Font(color="FFFFFF", bold=True)
    for col in list(range(1, len(EXCEL_HEADERS) + 1)) + [FILENAME_COL]:
        c = ws.cell(1, col)
        c.fill = fill
        c.font = font
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.freeze_panes = "A2"

    widths = {
        1: 10, 2: 14, 3: 18, 4: 16, 5: 28, 6: 12, 7: 28,
        8: 12, 9: 10, 10: 12, 11: 12, 12: 14, 13: 14, 14: 18,
        15: 12, 16: 14, 19: 28,
    }
    for col, width in widths.items():
        ws.column_dimensions[chr(64 + col) if col <= 26 else str(col)].width = width


def open_or_create_workbook(output: Path):
    if output.exists():
        wb = load_workbook(output)
        ws = wb["Data"] if "Data" in wb.sheetnames else wb.active
        ws.title = "Data"
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Data"
    ensure_excel_layout(ws)
    return wb, ws


def build_excel_index(ws) -> tuple[dict[str, int], dict[str, int]]:
    by_sno: dict[str, int] = {}
    by_fallback: dict[str, int] = {}
    for r in range(2, ws.max_row + 1):
        values = [clean_text(ws.cell(r, c).value) for c in range(1, 17)]
        sno = normalize_sno(values[0]).casefold()
        if sno and sno not in by_sno:
            by_sno[sno] = r
        fb = _fallback_identity(values)
        if any(values[:6]) and fb.strip("|") and fb not in by_fallback:
            by_fallback[fb] = r
    return by_sno, by_fallback


@dataclass
class UpsertResult:
    inserted: int = 0
    updated: int = 0
    date_conflicts: int = 0
    rows_touched: int = 0


def upsert_excel_records(output: Path, records: List[GKRecord], source_filename: str) -> UpsertResult:
    wb, ws = open_or_create_workbook(output)
    result = UpsertResult()
    try:
        by_sno, by_fallback = build_excel_index(ws)
        for rec in records:
            values = record_to_row(rec)
            sno_key = normalize_sno(rec.sno).casefold()
            fb_key = _fallback_identity(values)
            target = by_sno.get(sno_key) if sno_key else None
            if target is None and fb_key.strip("|"):
                target = by_fallback.get(fb_key)

            if target is None:
                target = max(ws.max_row + 1, 2)
                for c, value in enumerate(values, start=1):
                    ws.cell(target, c, value)
                ws.cell(target, FILENAME_COL, source_filename)
                result.inserted += 1
                if sno_key:
                    by_sno[sno_key] = target
                if fb_key.strip("|"):
                    by_fallback[fb_key] = target
            else:
                existing_date = clean_text(ws.cell(target, 2).value)
                incoming_date = clean_text(values[1])
                if existing_date and incoming_date and _date_norm(existing_date) != _date_norm(incoming_date):
                    values[1] = existing_date
                    result.date_conflicts += 1

                for c, value in enumerate(values, start=1):
                    # Never erase a good existing value because OCR returned blank.
                    if clean_text(value):
                        ws.cell(target, c, value)
                ws.cell(target, FILENAME_COL, source_filename)
                result.updated += 1
            result.rows_touched += 1

        wb.save(output)
        return result
    finally:
        wb.close()


# -----------------------------------------------------------------------------
# Google Apps Script writer
# -----------------------------------------------------------------------------


class AppsScriptGKWriter:
    def __init__(self, webapp_url: str, receiver_key: str, timeout: int = 60):
        self.webapp_url = clean_text(webapp_url)
        self.receiver_key = clean_text(receiver_key)
        self.timeout = timeout
        self.input_name = ""
        if not self.webapp_url.startswith("https://script.google.com/") or not self.webapp_url.endswith("/exec"):
            raise ValueError("SHEET_WEBAPP_URL must be the deployed Apps Script /exec URL")
        if not self.receiver_key:
            raise ValueError("SHEET_RECEIVER_KEY is empty")

    def _post(self, payload: dict) -> dict:
        body = {"key": self.receiver_key, "dataset": "gk", **payload}
        normal_attempt = 0
        lock_attempt = 0
        lock_delays = [10, 20, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30]
        while True:
            try:
                response = requests.post(self.webapp_url, json=body, timeout=self.timeout, allow_redirects=True)
                response.raise_for_status()
                result = response.json()
                if not result.get("ok"):
                    raise RuntimeError(result.get("error", "Apps Script rejected request"))
                return result
            except Exception as exc:
                text = str(exc).lower()
                if "lock timeout" in text or "holding the lock" in text:
                    if lock_attempt >= len(lock_delays):
                        raise RuntimeError(f"Google Sheet lock remained busy: {exc}") from exc
                    wait = lock_delays[lock_attempt]
                    lock_attempt += 1
                    time.sleep(wait)
                    continue
                normal_attempt += 1
                if normal_attempt < 3:
                    time.sleep(min(2 ** normal_attempt, 8))
                    continue
                raise RuntimeError(f"Google Sheet update failed after 3 attempts: {exc}") from exc

    def start(self, input_name: str, total_pages: int, run_id: str) -> None:
        self.input_name = input_name
        self._post({
            "action": "reset",
            "input_name": input_name,
            "total_pages": total_pages,
            "page_number": 0,
            "total_records": 0,
            "run_id": run_id,
        })

    def upsert_page(self, page_number: int, total_pages: int, records: List[GKRecord], total_records: int) -> dict:
        return self._post({
            "action": "upsert",
            "input_name": self.input_name,
            "page_number": page_number,
            "total_pages": total_pages,
            "total_records": total_records,
            "rows": [record_to_row(r) for r in records],
        })

    def status(self, state: str, page_number: int, total_pages: int, total_records: int, message: str) -> None:
        self._post({
            "action": "status",
            "input_name": self.input_name,
            "state": state,
            "page_number": page_number,
            "total_pages": total_pages,
            "total_records": total_records,
            "message": message,
        })

    def safe_status(self, **kwargs) -> None:
        try:
            self.status(**kwargs)
        except Exception:
            pass

    def finish(self, total_pages: int, total_records: int) -> None:
        self._post({
            "action": "finish",
            "input_name": self.input_name,
            "total_pages": total_pages,
            "total_records": total_records,
        })


# -----------------------------------------------------------------------------
# High-level file processing
# -----------------------------------------------------------------------------


@dataclass
class ProcessSummary:
    files: int = 0
    pages: int = 0
    extracted_records: int = 0
    excel_inserted: int = 0
    excel_updated: int = 0
    date_conflicts: int = 0


def process_pages(
    *,
    filename: str,
    pages,
    key_pool: GeminiKeyPool,
    model: str,
    excel_path: Optional[Path],
    sheet_writer: Optional[AppsScriptGKWriter],
    wait_on_quota: bool,
    usage: UsageTotals,
    log: Callable[[str], None] = print,
    progress_callback: Optional[Callable[[int, int, str], None]] = None,
) -> ProcessSummary:
    page_list = list(pages)
    total_pages = len(page_list)
    run_id = hashlib.sha256(filename.encode("utf-8") + b"|" + b"".join(p[2][:1024] for p in page_list)).hexdigest()
    summary = ProcessSummary(files=1)

    if sheet_writer:
        sheet_writer.start(filename, total_pages, run_id)

    total_records = 0
    for page_number, _, payload, mime_type, _page_id in page_list:
        if progress_callback:
            progress_callback(page_number, total_pages, f"Reading {filename} - page {page_number}/{total_pages}")
        log(f"{filename}: page {page_number}/{total_pages} - Gemini extraction")

        def status_cb(state: str, message: str) -> None:
            if sheet_writer:
                sheet_writer.safe_status(
                    state=state,
                    page_number=page_number,
                    total_pages=total_pages,
                    total_records=total_records,
                    message=message,
                )

        records = extract_page(
            key_pool=key_pool,
            media_bytes=payload,
            mime_type=mime_type,
            model=model,
            media_resolution="high",
            thinking_level="low",
            retries=3,
            usage=usage,
            wait_on_quota=wait_on_quota,
            log=log,
            status_callback=status_cb,
        )
        total_records += len(records)
        summary.pages += 1
        summary.extracted_records += len(records)

        if excel_path is not None:
            x = upsert_excel_records(excel_path, records, filename)
            summary.excel_inserted += x.inserted
            summary.excel_updated += x.updated
            summary.date_conflicts += x.date_conflicts
            log(
                f"  Excel: inserted={x.inserted}, updated={x.updated}, "
                f"date-conflict-protected={x.date_conflicts}"
            )

        if sheet_writer:
            result = sheet_writer.upsert_page(page_number, total_pages, records, total_records)
            log(
                "  Google Sheet: "
                f"inserted={result.get('inserted', 0)}, updated={result.get('updated', 0)}, "
                f"date conflicts={result.get('date_conflicts', 0)}"
            )

    if sheet_writer:
        sheet_writer.finish(total_pages, total_records)
    return summary
