#!/usr/bin/env python3
"""Streamlit uploader for GK register images/PDFs -> Google Sheet."""

from __future__ import annotations

import tempfile
import tomllib
from pathlib import Path

import streamlit as st

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


def _load_toml_secrets() -> dict:
    """
    Fallback secrets loader.

    Tries:
      1) .streamlit/secrets.toml beside this script
      2) .streamlit/secrets.toml in the current working directory
    """
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
    """
    Read a setting from Streamlit secrets first.
    If Streamlit does not load secrets correctly, read the TOML file directly.
    """
    # 1. Normal Streamlit secrets.
    try:
        value = st.secrets.get(name, "")
        if value is not None and str(value).strip():
            return str(value).strip()
    except Exception:
        pass

    # 2. Direct TOML fallback.
    try:
        value = TOML_SECRETS.get(name, "")
        if value is not None and str(value).strip():
            return str(value).strip()
    except Exception:
        pass

    return default


# Optional password gate.
APP_PASSWORD = secret("APP_PASSWORD")
if APP_PASSWORD:
    entered = st.text_input("Access password", type="password")
    if entered != APP_PASSWORD:
        st.info("Enter the access password to continue.")
        st.stop()


st.title("GK Register → Google Sheet")
st.caption(
    "Upload register photos/scans or PDFs. Existing rows are updated by S.No; "
    "duplicate rows are not blindly appended."
)



api1 = secret("GEMINI_API_KEY_1")
api2 = secret("GEMINI_API_KEY_2")
webapp = secret("SHEET_WEBAPP_URL")
receiver = secret("SHEET_RECEIVER_KEY")
model = secret("GEMINI_MODEL", "gemini-3.5-flash")

missing = [
    n
    for n, v in {
        "GEMINI_API_KEY_1": api1,
        "SHEET_WEBAPP_URL": webapp,
        "SHEET_RECEIVER_KEY": receiver,
    }.items()
    if not str(v).strip()
]

if missing:
    st.error("Missing Streamlit secrets: " + ", ".join(missing))

    # Safe diagnostics: never print secret values.
    script_secrets = Path(__file__).resolve().parent / ".streamlit" / "secrets.toml"
    cwd_secrets = Path.cwd() / ".streamlit" / "secrets.toml"

    st.code(
        "\n".join(
            [
                f"Script folder: {Path(__file__).resolve().parent}",
                f"Current working folder: {Path.cwd()}",
                f"Script-folder secrets.toml exists: {script_secrets.exists()}",
                f"CWD secrets.toml exists: {cwd_secrets.exists()}",
                f"GEMINI_API_KEY_1 loaded: {'YES' if api1 else 'NO'}",
                f"GEMINI_API_KEY_2 loaded: {'YES' if api2 else 'NO'}",
                f"SHEET_WEBAPP_URL loaded: {'YES' if webapp else 'NO'}",
                f"SHEET_RECEIVER_KEY loaded: {'YES' if receiver else 'NO'}",
            ]
        )
    )
    st.stop()


uploads = st.file_uploader(
    "Upload image(s) or PDF(s)",
    type=["jpg", "jpeg", "png", "webp", "heic", "heif", "tif", "tiff", "pdf"],
    accept_multiple_files=True,
)

if uploads:
    cols = st.columns(min(3, len(uploads)))
    for i, f in enumerate(uploads[:3]):
        if Path(f.name).suffix.lower() != ".pdf":
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
    key_pool = GeminiKeyPool([api1, api2])
    writer = AppsScriptGKWriter(webapp, receiver)
    usage = UsageTotals()

    progress = st.progress(0.0)
    status = st.empty()
    log_box = st.empty()
    log_lines: list[str] = []

    def log(msg: str) -> None:
        log_lines.append(str(msg))
        log_box.code("\n".join(log_lines[-30:]))

    with tempfile.TemporaryDirectory() as td:
        session_excel = Path(td) / "gk_register_session.xlsx"
        overall_files = len(uploads)
        done_files = 0

        totals = {
            "records": 0,
            "inserted": 0,
            "updated": 0,
            "conflicts": 0,
        }

        try:
            for uploaded in uploads:
                data = uploaded.getvalue()
                pages = list(iter_uploaded_pages(uploaded.name, data))

                def on_progress(
                    page_no: int,
                    total_pages: int,
                    message: str,
                ) -> None:
                    frac = (
                        done_files
                        + (page_no - 1) / max(total_pages, 1)
                    ) / max(overall_files, 1)

                    progress.progress(min(0.99, frac))
                    status.info(message)

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
                progress.progress(done_files / overall_files)

            status.success("Finished. Google Sheet has been updated.")

            st.success(
                f"Extracted {totals['records']} row(s). Session Excel: "
                f"{totals['inserted']} inserted, "
                f"{totals['updated']} updated, "
                f"{totals['conflicts']} date conflict(s) protected."
            )

            st.caption(
                f"Gemini tokens this session: "
                f"input {usage.input_tokens}, "
                f"output {usage.output_tokens}, "
                f"thinking {usage.thinking_tokens}, "
                f"total {usage.total_tokens}."
            )

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
            status.warning(str(exc))
            st.warning(
                "Both configured Gemini projects are currently quota/rate limited. "
                "Try again after the quota resets."
            )

        except Exception as exc:
            status.error("Processing stopped")
            st.exception(exc)
