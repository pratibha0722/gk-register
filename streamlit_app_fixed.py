#!/usr/bin/env python3
"""GK Register Streamlit uploader with Google Drive persistence + background queue."""

from __future__ import annotations

import base64
import hashlib
import mimetypes
import time
import tomllib
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import requests
import streamlit as st

IST = ZoneInfo("Asia/Kolkata")

st.set_page_config(
    page_title="GK Register Upload",
    page_icon="📄",
    layout="wide",
)


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

    value = TOML_SECRETS.get(name, "")
    if value is not None and str(value).strip():
        return str(value).strip()
    return default


WEBAPP_URL = secret("SHEET_WEBAPP_URL")
RECEIVER_KEY = secret("SHEET_RECEIVER_KEY")
APP_PASSWORD = secret("APP_PASSWORD")

missing = [
    name
    for name, value in {
        "SHEET_WEBAPP_URL": WEBAPP_URL,
        "SHEET_RECEIVER_KEY": RECEIVER_KEY,
    }.items()
    if not value
]

if missing:
    st.error("Missing Streamlit secrets: " + ", ".join(missing))
    st.stop()

if APP_PASSWORD:
    entered = st.text_input("Access password", type="password")
    if entered != APP_PASSWORD:
        st.info("Enter the access password to continue.")
        st.stop()


def receiver_post(payload: dict, timeout: int = 90) -> dict:
    body = {
        "key": RECEIVER_KEY,
        "dataset": "gk",
        **payload,
    }
    response = requests.post(
        WEBAPP_URL,
        json=body,
        timeout=timeout,
        allow_redirects=True,
    )
    response.raise_for_status()
    result = response.json()
    if not result.get("ok"):
        raise RuntimeError(result.get("error", "Receiver rejected request"))
    return result


def mime_for(name: str, uploaded_type: str | None) -> str:
    if uploaded_type:
        return uploaded_type
    guessed, _ = mimetypes.guess_type(name)
    return guessed or "application/octet-stream"


st.title("GK Register → Google Sheet")
st.caption(
    "Upload images/PDFs or take a photo. Files are saved to your restricted "
    "Google Drive folder first, then processed from a persistent queue."
)

uploads = st.file_uploader(
    "Upload image(s) or PDF(s)",
    type=["jpg", "jpeg", "png", "webp", "heic", "heif", "tif", "tiff", "pdf"],
    accept_multiple_files=True,
)

# Camera permission is requested only after the user explicitly clicks Take photo.
if "show_camera" not in st.session_state:
    st.session_state.show_camera = False

if st.button("📷 Take photo"):
    st.session_state.show_camera = True

camera_photo = None
if st.session_state.show_camera:
    camera_photo = st.camera_input("Camera")

all_uploads = list(uploads or [])
if camera_photo is not None:
    all_uploads.append(camera_photo)

if all_uploads:
    images = [f for f in all_uploads if Path(f.name).suffix.lower() != ".pdf"]
    if images:
        cols = st.columns(min(3, len(images)))
        for i, f in enumerate(images[:3]):
            try:
                cols[i % len(cols)].image(
                    f.getvalue(),
                    caption=f.name,
                    use_container_width=True,
                )
            except Exception:
                pass


if "last_queue_ids" not in st.session_state:
    st.session_state.last_queue_ids = []
if "history_items" not in st.session_state:
    st.session_state.history_items = []
if "history_queue_ids" not in st.session_state:
    st.session_state.history_queue_ids = []
if "history_loaded_date" not in st.session_state:
    st.session_state.history_loaded_date = ""
if "history_worker_active" not in st.session_state:
    st.session_state.history_worker_active = True

if st.button(
    "Upload to Drive and start processing",
    type="primary",
    disabled=not all_uploads,
):
    progress = st.progress(0.0)
    status = st.empty()
    queued = []
    total = len(all_uploads)

    try:
        for i, uploaded in enumerate(all_uploads, start=1):
            raw = uploaded.getvalue()
            display_name = uploaded.name

            # Camera widgets often return a generic filename. Give camera images
            # a useful unique name before saving them to Drive.
            if uploaded is camera_photo:
                suffix = Path(display_name).suffix.lower() or ".jpg"
                display_name = datetime.now(IST).strftime("camera_%Y%m%d_%H%M%S") + suffix

            status.info(f"Saving {i}/{total}: {display_name}")
            sha256 = hashlib.sha256(raw).hexdigest()

            result = receiver_post(
                {
                    "action": "enqueue_upload",
                    "file_name": display_name,
                    "mime_type": mime_for(display_name, getattr(uploaded, "type", None)),
                    "content_sha256": sha256,
                    "data_base64": base64.b64encode(raw).decode("ascii"),
                },
                timeout=120,
            )

            queued.append(result)
            progress.progress(i / total)

        st.session_state.last_queue_ids = [
            str(item.get("queue_id", ""))
            for item in queued
            if item.get("queue_id")
        ]

        status.success(
            f"{len(queued)} file(s) saved to Drive and queued successfully."
        )
        st.success(
            "You can close this page now. The background worker continues from "
            "Google Drive and automatically retries pending files when API quota "
            "becomes available again."
        )

        for item in queued:
            name = item.get("file_name", "File")
            drive_url = item.get("drive_url", "")
            queue_status = item.get("status", "Pending")
            reason = item.get("reason", "Waiting for background worker")
            if drive_url:
                st.markdown(
                    f"- [{name}]({drive_url}) — **{queue_status}** — {reason}"
                )
            else:
                st.write(f"- {name} — {queue_status} — {reason}")

    except Exception as exc:
        status.error("Upload/queue operation stopped.")
        st.exception(exc)


if st.session_state.last_queue_ids:
    col_refresh, col_retry = st.columns(2)

    refresh_clicked = col_refresh.button("Refresh processing status")
    retry_clicked = col_retry.button("Retry waiting files now")

    if retry_clicked:
        try:
            retry_result = receiver_post(
                {
                    "action": "queue_retry_now",
                    "queue_ids": st.session_state.last_queue_ids,
                },
                timeout=60,
            )
            changed = int(retry_result.get("changed", 0) or 0)
            if changed:
                st.success(
                    f"{changed} waiting file(s) re-queued. No re-upload was needed."
                )
            else:
                st.info("No waiting/error files needed a manual retry.")
            refresh_clicked = True
        except Exception as exc:
            st.exception(exc)

    if refresh_clicked:
        try:
            result = receiver_post(
                {
                    "action": "queue_status",
                    "queue_ids": st.session_state.last_queue_ids,
                },
                timeout=60,
            )
            items = result.get("items", [])
            worker_trigger_active = result.get("worker_trigger_active", True)

            if not worker_trigger_active:
                st.error(
                    "Background worker trigger is missing. In Apps Script, run "
                    "repairGKWorkerTrigger once. Existing Drive files do not need "
                    "to be uploaded again."
                )

            if items:
                completed = sum(1 for x in items if x.get("status") == "Completed")
                processing = sum(1 for x in items if x.get("status") == "Processing")
                waiting = len(items) - completed - processing
                st.caption(
                    f"This upload session: {completed} completed • "
                    f"{processing} processing • {waiting} waiting/retrying"
                )

                rows = []
                for item in items:
                    status_value = item.get("status", "")
                    rows.append(
                        {
                            "File": item.get("file_name", ""),
                            "Status": status_value,
                            "Reason": item.get("reason", item.get("last_error", "")),
                            "Attempts": item.get("attempts", 0),
                            "Records": item.get("records", 0),
                            "Last update": item.get("last_update", ""),
                            "Eligible after": item.get("next_attempt", ""),
                        }
                    )
                st.dataframe(rows, use_container_width=True, hide_index=True)

                for item in items:
                    name = item.get("file_name", "File")
                    status_value = item.get("status", "")
                    records = int(item.get("records", 0) or 0)
                    reason = item.get("reason", "")
                    if status_value == "Completed":
                        st.success(f"✅ {name}: completed ({records} records). {reason}")
                    elif status_value == "Processing":
                        st.info(f"⏳ {name}: processing now.")
                    elif status_value in {"Retry", "Waiting for quota"}:
                        st.warning(f"⚠️ {name}: {status_value}. {reason}")
                    elif status_value == "Error":
                        st.error(f"❌ {name}: {reason}")
            else:
                st.info("No queue records found for this upload session.")
        except Exception as exc:
            st.exception(exc)

st.divider()
st.subheader("Previous uploads / daily history")
st.caption(
    "Use this after a page refresh, browser restart, or on a later day to check "
    "whether files were saved to Drive and whether processing completed."
)

history_date = st.date_input(
    "Choose upload date",
    value=datetime.now(IST).date(),
    key="history_date_selector",
)
history_date_iso = history_date.isoformat()

history_col_check, history_col_retry = st.columns(2)
history_check_clicked = history_col_check.button(
    "Check uploads for selected day",
    use_container_width=True,
)
history_retry_clicked = history_col_retry.button(
    "Retry waiting files from selected day",
    use_container_width=True,
    disabled=not bool(st.session_state.history_queue_ids)
    or st.session_state.history_loaded_date != history_date_iso,
)


def load_history_for_day(day_iso: str) -> None:
    result = receiver_post(
        {
            "action": "queue_status",
            "date_ist": day_iso,
        },
        timeout=60,
    )
    items = result.get("items", [])
    st.session_state.history_items = items
    st.session_state.history_queue_ids = [
        str(item.get("queue_id", ""))
        for item in items
        if item.get("queue_id")
    ]
    st.session_state.history_loaded_date = day_iso
    st.session_state.history_worker_active = result.get(
        "worker_trigger_active", True
    )


if history_check_clicked:
    try:
        load_history_for_day(history_date_iso)
    except Exception as exc:
        st.exception(exc)

if history_retry_clicked:
    try:
        retry_result = receiver_post(
            {
                "action": "queue_retry_now",
                "queue_ids": st.session_state.history_queue_ids,
            },
            timeout=60,
        )
        changed = int(retry_result.get("changed", 0) or 0)
        if changed:
            st.success(
                f"{changed} waiting/error file(s) from this day were re-queued. "
                "No re-upload was needed."
            )
        else:
            st.info("No waiting/error files from this day needed a manual retry.")
        load_history_for_day(history_date_iso)
    except Exception as exc:
        st.exception(exc)

if st.session_state.history_loaded_date == history_date_iso:
    history_items = st.session_state.history_items

    if not st.session_state.history_worker_active:
        st.error(
            "Background worker trigger is missing. In Apps Script run "
            "repairGKWorkerTrigger once. Files already in Drive do not need "
            "to be uploaded again."
        )

    if history_items:
        completed = sum(
            1 for item in history_items if item.get("status") == "Completed"
        )
        processing = sum(
            1 for item in history_items if item.get("status") == "Processing"
        )
        waiting = len(history_items) - completed - processing

        st.caption(
            f"Selected day: {len(history_items)} file(s) uploaded to Drive | "
            f"{completed} completed | {processing} processing | "
            f"{waiting} waiting/retrying/error"
        )

        history_rows = []
        for item in history_items:
            history_rows.append(
                {
                    "File": item.get("file_name", ""),
                    "Drive saved": "Yes" if item.get("drive_url") else "No",
                    "Status": item.get("status", ""),
                    "Reason": item.get(
                        "reason", item.get("last_error", "")
                    ),
                    "Attempts": item.get("attempts", 0),
                    "Records": item.get("records", 0),
                    "Uploaded at": item.get("added_at", ""),
                    "Last update": item.get("last_update", ""),
                    "Eligible after": item.get("next_attempt", ""),
                }
            )

        st.dataframe(
            history_rows,
            use_container_width=True,
            hide_index=True,
        )

        st.caption("Open the stored Drive image/file:")
        for item in history_items:
            drive_url = item.get("drive_url", "")
            if not drive_url:
                continue
            name = str(item.get("file_name", "File")).replace("[", "\\[").replace("]", "\\]")
            status_value = item.get("status", "")
            st.markdown(f"- [{name}]({drive_url}) - **{status_value}**")
    else:
        st.info("No uploads were found for the selected date.")
elif st.session_state.history_loaded_date:
    st.info("Date changed. Click 'Check uploads for selected day' to load that date.")

