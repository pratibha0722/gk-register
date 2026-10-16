#!/usr/bin/env python3
"""GK Register Streamlit app with clear live processing status."""

from __future__ import annotations

import tempfile
import time
import tomllib
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import streamlit as st
IST =ZoneInfo("Asia/Kolkata")
from gk_core import (
    AppsScriptGKWriter,
    GeminiKeyPool,
    QuotaUnavailable,
    UsageTotals,
    iter_uploaded_pages,
    process_pages,
)

st.set_page_config(
    page_title="GK Register Upload",
    page_icon="📄",
    layout="wide",
)


# ---------------------------------------------------------------------
# Secrets
# ---------------------------------------------------------------------

def _load_toml_secrets() -> dict:
    candidates = [
        Path(__file__).resolve().parent / ".streamlit" / "secrets.toml",
        Path.cwd() / ".streamlit" / "secrets.toml",
    ]
    for path in candidates:
        try:
            if path.exists():
                with path.open("rb") as fh:
                    data = tomllib.load(fh)
                if isinstance(data, dict):
                    return data
        except Exception:
            pass
    return {}


TOML_SECRETS = _load_toml_secrets()


def secret(name: str, default: str = "") -> str:
    try:
        value = st.secrets.get(name, "")
        if value is not None and str(value).strip():
            return str(value).strip()
    except Exception:
        pass

    try:
        value = TOML_SECRETS.get(name, "")
        if value is not None and str(value).strip():
            return str(value).strip()
    except Exception:
        pass

    return default


api1 = secret("GEMINI_API_KEY_1")
api2 = secret("GEMINI_API_KEY_2")
webapp = secret("SHEET_WEBAPP_URL")
receiver = secret("SHEET_RECEIVER_KEY")
model = secret("GEMINI_MODEL", "gemini-3.5-flash")

missing = [
    name
    for name, value in {
        "GEMINI_API_KEY_1": api1,
        "SHEET_WEBAPP_URL": webapp,
        "SHEET_RECEIVER_KEY": receiver,
    }.items()
    if not str(value).strip()
]

if missing:
    st.error("Missing Streamlit secrets: " + ", ".join(missing))
    st.stop()


# ---------------------------------------------------------------------
# Optional password
# ---------------------------------------------------------------------

APP_PASSWORD = secret("APP_PASSWORD")
if APP_PASSWORD:
    entered = st.text_input("Access password", type="password")
    if entered != APP_PASSWORD:
        st.info("Enter the access password to continue.")
        st.stop()


# ---------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------

st.title("GK Register → Google Sheet")
st.caption("Upload register photos/scans or PDFs and update the Google Sheet.")

uploads = st.file_uploader(
    "Upload image(s) or PDF(s)",
    type=["jpg", "jpeg", "png", "webp", "heic", "heif", "tif", "tiff", "pdf"],
    accept_multiple_files=True,
)

if uploads:
    image_uploads = [
        f for f in uploads
        if Path(f.name).suffix.lower() != ".pdf"
    ]
    if image_uploads:
        cols = st.columns(min(3, len(image_uploads)))
        for i, f in enumerate(image_uploads[:3]):
            try:
                cols[i % len(cols)].image(
                    f.getvalue(),
                    caption=f.name,
                    use_container_width=True,
                )
            except Exception:
                pass


if st.button(
    "Process and update Google Sheet",
    type="primary",
    disabled=not uploads,
):
    # This app processes synchronously. Keep the browser session open.
    st.warning(
        "Keep this browser tab open until the job says Finished. "
        "This Streamlit version processes the upload in the active web session."
    )

    job_state = st.empty()
    current_state = st.empty()
    activity_state = st.empty()
    progress = st.progress(0.0)
    log_box = st.empty()

    started_monotonic = time.monotonic()
    started_at = datetime.now(IST)

    job_state.info(
        f"🟢 PROCESSING ACTIVE — started {started_at.strftime('%H:%M:%S')}"
    )

    log_lines: list[str] = []

    def elapsed_text() -> str:
        sec = int(time.monotonic() - started_monotonic)
        mins, seconds = divmod(sec, 60)
        hours, mins = divmod(mins, 60)
        if hours:
            return f"{hours}h {mins}m {seconds}s"
        if mins:
            return f"{mins}m {seconds}s"
        return f"{seconds}s"

    def touch(message: str) -> None:
        activity_state.caption(
            f"Last activity: {datetime.now(IST).strftime('%H:%M:%S')} "
            f"• elapsed {elapsed_text()} • {message}"
        )

    def log(message: str) -> None:
        stamped = f"[{datetime.now(IST).strftime('%H:%M:%S')}] {message}"
        log_lines.append(stamped)
        log_box.code("\n".join(log_lines[-40:]))
        touch(str(message))

    key_pool = GeminiKeyPool([api1, api2])
    writer = AppsScriptGKWriter(webapp, receiver)
    usage = UsageTotals()

    with tempfile.TemporaryDirectory() as td:
        session_excel = Path(td) / "gk_register_session.xlsx"

        total_files = len(uploads)
        done_files = 0

        totals = {
            "records": 0,
            "inserted": 0,
            "updated": 0,
            "conflicts": 0,
        }

        try:
            for file_index, uploaded in enumerate(uploads, start=1):
                data = uploaded.getvalue()
                pages = list(iter_uploaded_pages(uploaded.name, data))
                total_pages = max(1, len(pages))

                current_state.info(
                    f"File {file_index}/{total_files}: **{uploaded.name}** "
                    f"• {total_pages} page(s)"
                )
                touch(f"Preparing {uploaded.name}")

                def on_progress(
                    page_no: int,
                    file_total_pages: int,
                    message: str,
                ) -> None:
                    # Show some visible progress while the current page is being read.
                    page_fraction = (
                        (page_no - 1) / max(file_total_pages, 1)
                    )
                    overall_fraction = (
                        done_files + page_fraction
                    ) / max(total_files, 1)

                    progress.progress(
                        min(0.99, max(0.01, overall_fraction))
                    )

                    current_state.info(
                        f"File {file_index}/{total_files}: **{uploaded.name}** "
                        f"• page {page_no}/{file_total_pages} "
                        f"• LLM processing"
                    )
                    touch(message)

                # The animated spinner remains visible even while the LLM
                # request itself is blocking and there are no new log messages.
                with st.spinner(
                    f"LLM is processing {uploaded.name}. "
                    "The animated spinner means this request is still active..."
                ):
                    summary = process_pages(
                        filename=uploaded.name,
                        pages=pages,
                        key_pool=key_pool,
                        model=model,
                        excel_path=session_excel,
                        sheet_writer=writer,
                        wait_on_quota=False,
                        usage=usage,
                        log=log,
                        progress_callback=on_progress,
                    )

                totals["records"] += summary.extracted_records
                totals["inserted"] += summary.excel_inserted
                totals["updated"] += summary.excel_updated
                totals["conflicts"] += summary.date_conflicts

                done_files += 1
                progress.progress(done_files / total_files)

                log(
                    f"Completed {uploaded.name}: "
                    f"{summary.extracted_records} row(s)"
                )

            current_state.success("All uploaded files processed.")
            touch("Job completed")
            job_state.success(
                f"✅ FINISHED — elapsed {elapsed_text()}"
            )

            st.success(
                f"Extracted {totals['records']} row(s). "
                f"Session Excel: {totals['inserted']} inserted, "
                f"{totals['updated']} updated, "
                f"{totals['conflicts']} date conflict(s) protected."
            )

            st.caption(
                f"LLM tokens this session: "
                f"input {usage.input_tokens}, "
                f"output {usage.output_tokens}, "
                f"thinking {usage.thinking_tokens}, "
                f"total {usage.total_tokens}."
            )

            if session_excel.exists():
                with open(session_excel, "rb") as fh:
                    st.download_button(
                        "Download this session Excel",
                        data=fh.read(),
                        file_name="gk_register_session.xlsx",
                        mime=(
                            "application/vnd.openxmlformats-officedocument."
                            "spreadsheetml.sheet"
                        ),
                    )

        except QuotaUnavailable as exc:
            job_state.warning("🟡 PAUSED / QUOTA UNAVAILABLE")
            current_state.warning(str(exc))
            touch("LLM quota/rate limit reached")
            st.warning(
                "Both configured LLM projects are currently quota/rate "
                "limited. No further files were processed."
            )

        except Exception as exc:
            job_state.error("🔴 STOPPED WITH ERROR")
            current_state.error("Processing stopped")
            touch("Processing stopped with an error")
            st.exception(exc)
