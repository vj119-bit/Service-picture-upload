import mimetypes
import os
import re
import json
import hashlib
import shutil
import uuid
from datetime import datetime
from pathlib import Path
from urllib.parse import quote

import requests
import streamlit as st
import streamlit.components.v1 as components
from msal import ConfidentialClientApplication


# ============================================================
# CONFIG LOADER
# Priority:
# 1) Streamlit secrets
# 2) Environment variables
# ============================================================

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
SMALL_UPLOAD_LIMIT = 250 * 1024 * 1024  # 250 MB
CHUNK_SIZE = 327680 * 20  # 6.25 MB, valid multiple of 320 KiB
MAX_PER_FILE_UPLOAD_MB = 1000
MAX_TOTAL_UPLOAD_MB = 1000
MAX_FILES_PER_BATCH = 15
RECOMMENDED_MOBILE_TOTAL_MB = 300
STAGING_ROOT = Path(".upload_staging")


def is_placeholder(value: str) -> bool:
    raw = (value or "").strip()
    lowered = raw.lower()
    return (
        not raw
        or "paste" in lowered
        or "yourtenant" in lowered
        or "yoursitename" in lowered
    )


def secret_or_env(section: str, key: str, env_key: str = None, default: str = "") -> str:
    env_key = env_key or key.upper()
    try:
        if section in st.secrets and key in st.secrets[section]:
            section_value = str(st.secrets[section][key])
            if not is_placeholder(section_value):
                return section_value
        if env_key in st.secrets:
            flat_value = str(st.secrets[env_key])
            if not is_placeholder(flat_value):
                return flat_value
    except Exception:
        pass
    return os.getenv(env_key, default)


def normalize_sharepoint_hostname(value: str) -> str:
    value = (value or "").strip()
    value = re.sub(r"^https?://", "", value, flags=re.IGNORECASE)
    return value.rstrip("/")


TENANT_ID = secret_or_env("azure", "tenant_id", "TENANT_ID")
CLIENT_ID = secret_or_env("azure", "client_id", "CLIENT_ID")
CLIENT_SECRET = secret_or_env("azure", "client_secret", "CLIENT_SECRET")

SHAREPOINT_HOSTNAME = normalize_sharepoint_hostname(
    secret_or_env("sharepoint", "hostname", "SHAREPOINT_HOSTNAME")
)
SHAREPOINT_SITE_PATH = secret_or_env("sharepoint", "site_path", "SHAREPOINT_SITE_PATH")
DOCUMENT_LIBRARY_NAME = secret_or_env("sharepoint", "document_library_name", "DOCUMENT_LIBRARY_NAME", "Documents")
BASE_FOLDER_NAME = secret_or_env("sharepoint", "base_folder_name", "BASE_FOLDER_NAME", "Customer Uploads")

APP_TITLE = secret_or_env("branding", "app_title", "APP_TITLE", "Customer Media Upload Portal")
APP_SUBTITLE = secret_or_env(
    "branding",
    "app_subtitle",
    "APP_SUBTITLE",
    "Upload customer photos and videos directly to SharePoint."
)
COMPANY_NAME = secret_or_env("branding", "company_name", "COMPANY_NAME", "Your Company")

EMAIL_NOTIFY_ENABLED = secret_or_env("notification", "enabled", "EMAIL_NOTIFY_ENABLED", "false").lower() == "true"
NOTIFY_SENDER_EMAIL = secret_or_env("notification", "sender_email", "NOTIFY_SENDER_EMAIL")
NOTIFY_TO_EMAIL = secret_or_env("notification", "to_email", "NOTIFY_TO_EMAIL")


# ============================================================
# PAGE / THEME
# ============================================================

st.set_page_config(page_title=APP_TITLE, page_icon="📤", layout="centered")

st.markdown(
    """
    <style>
    .main > div {
        padding-top: 1.5rem;
    }
    .hero-card {
        border: 1px solid rgba(49, 51, 63, 0.2);
        border-radius: 18px;
        padding: 1.25rem 1.25rem 1rem 1.25rem;
        background: linear-gradient(135deg, rgba(0, 104, 201, 0.08), rgba(0, 104, 201, 0.02));
        margin-bottom: 1rem;
    }
    .brand-kicker {
        font-size: 0.85rem;
        text-transform: uppercase;
        letter-spacing: 0.08em;
        opacity: 0.75;
        margin-bottom: 0.35rem;
    }
    .hero-title {
        font-size: 1.8rem;
        font-weight: 700;
        margin-bottom: 0.3rem;
    }
    .hero-subtitle {
        font-size: 1rem;
        opacity: 0.85;
        line-height: 1.45;
    }
    .hint-box {
        border-left: 4px solid rgba(0, 104, 201, 0.85);
        padding: 0.7rem 0.9rem;
        border-radius: 0.5rem;
        background: rgba(0, 104, 201, 0.05);
        margin-bottom: 1rem;
    }
    .footer-note {
        font-size: 0.85rem;
        opacity: 0.72;
        margin-top: 1rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    f"""
    <div class="hero-card">
        <div class="brand-kicker">{COMPANY_NAME}</div>
        <div class="hero-title">{APP_TITLE}</div>
        <div class="hero-subtitle">{APP_SUBTITLE}</div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="hint-box">
        Enter the customer details, choose one or more photos/videos, and press <b>Upload</b>.
        The app will create the customer folder automatically in SharePoint.
    </div>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# HELPERS
# ============================================================

def sanitize_name(value: str) -> str:
    """Make safe SharePoint/OneDrive file and folder names."""
    value = (value or "").strip()
    value = re.sub(r'[\"*:<>?/\\\\|]', "-", value)
    value = value.strip(" .")
    value = re.sub(r"\s+", " ", value)
    return value or "Unknown Customer"


def looks_like_guid(value: str) -> bool:
    return bool(re.fullmatch(r"[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}", (value or "").strip()))


def validate_config():
    missing = []
    fields = {
        "TENANT_ID": TENANT_ID,
        "CLIENT_ID": CLIENT_ID,
        "CLIENT_SECRET": CLIENT_SECRET,
        "SHAREPOINT_HOSTNAME": SHAREPOINT_HOSTNAME,
        "SHAREPOINT_SITE_PATH": SHAREPOINT_SITE_PATH,
    }

    for name, value in fields.items():
        if is_placeholder(value):
            missing.append(name)

    if CLIENT_SECRET and looks_like_guid(CLIENT_SECRET):
        st.error("Invalid CLIENT_SECRET format.")
        st.write("It looks like a Secret ID (GUID), not a Secret Value.")
        st.write("In Microsoft Entra > App registrations > your app > Certificates & secrets, create a new client secret and copy the Value field.")
        st.write("Do not use Secret ID. Using Secret ID causes AADSTS7000215.")
        st.stop()

    if missing:
        st.error("Configuration is incomplete.")
        st.write("Fill these values in `.streamlit/secrets.toml` or Streamlit app secrets.")
        st.write("Note: `.streamlit/secrets.toml.example` is only a sample file and is not loaded by Streamlit.")
        for item in missing:
            st.write(f"- {item}")
        st.stop()


@st.cache_resource
def get_msal_app() -> ConfidentialClientApplication:
    return ConfidentialClientApplication(
        client_id=CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        client_credential=CLIENT_SECRET,
    )


def get_access_token() -> str:
    result = get_msal_app().acquire_token_for_client(
        scopes=["https://graph.microsoft.com/.default"]
    )
    if "access_token" not in result:
        raise RuntimeError(
            f"Token error: {result.get('error')} - {result.get('error_description')}"
        )
    return result["access_token"]


def graph_request(method: str, url_or_path: str, allow_404: bool = False, **kwargs):
    url = url_or_path if url_or_path.startswith("http") else f"{GRAPH_BASE}{url_or_path}"

    headers = kwargs.pop("headers", {})
    headers["Authorization"] = f"Bearer {get_access_token()}"

    response = requests.request(method, url, headers=headers, timeout=180, **kwargs)

    if allow_404 and response.status_code == 404:
        return None

    if response.status_code >= 400:
        try:
            detail = response.json()
        except Exception:
            detail = response.text
        raise RuntimeError(f"Graph error {response.status_code}: {detail}")

    return response


def graph_json(method: str, url_or_path: str, allow_404: bool = False, **kwargs):
    response = graph_request(method, url_or_path, allow_404=allow_404, **kwargs)
    if response is None:
        return None
    if not response.content:
        return {}
    return response.json()


def get_site_id() -> str:
    site_path = SHAREPOINT_SITE_PATH.strip()
    if not site_path.startswith("/"):
        site_path = "/" + site_path

    encoded_site_path = quote(site_path, safe="/")
    data = graph_json("GET", f"/sites/{SHAREPOINT_HOSTNAME}:{encoded_site_path}")
    return data["id"]


def get_drive_id(site_id: str) -> str:
    target_library = DOCUMENT_LIBRARY_NAME.strip().lower()
    drives = graph_json("GET", f"/sites/{site_id}/drives").get("value", [])

    if not drives:
        raise RuntimeError("No document libraries found in the SharePoint site.")

    for drive in drives:
        if drive.get("name", "").strip().lower() == target_library:
            return drive["id"]

    available = ", ".join(d.get("name", "(unnamed)") for d in drives)
    raise RuntimeError(
        f"Document library '{DOCUMENT_LIBRARY_NAME}' not found. Available libraries: {available}"
    )


def get_root_item(drive_id: str) -> dict:
    return graph_json("GET", f"/drives/{drive_id}/root")


def get_item_by_path(drive_id: str, item_path: str):
    item_path = item_path.strip("/")
    if not item_path:
        return get_root_item(drive_id)

    encoded = quote(item_path, safe="/")
    return graph_json("GET", f"/drives/{drive_id}/root:/{encoded}", allow_404=True)


def ensure_folder_path(drive_id: str, folder_path: str) -> dict:
    parts = [p for p in folder_path.strip("/").split("/") if p]
    if not parts:
        return get_root_item(drive_id)

    current_path = []
    parent_item = get_root_item(drive_id)

    for part in parts:
        current_path.append(part)
        path_so_far = "/".join(current_path)

        existing = get_item_by_path(drive_id, path_so_far)
        if existing:
            parent_item = existing
            continue

        payload = {
            "name": part,
            "folder": {},
            "@microsoft.graph.conflictBehavior": "fail",
        }

        parent_item = graph_json(
            "POST",
            f"/drives/{drive_id}/items/{parent_item['id']}/children",
            json=payload,
        )

    return parent_item


def guess_extension(filename: str, content_type: str) -> str:
    ext = Path(filename).suffix.lower()
    if ext:
        return ext
    guessed = mimetypes.guess_extension(content_type or "")
    return guessed or ""


def format_size(num_bytes: int) -> str:
    units = ["B", "KB", "MB", "GB"]
    value = float(num_bytes)
    for unit in units:
        if value < 1024 or unit == units[-1]:
            return f"{value:.1f} {unit}" if unit != "B" else f"{int(value)} {unit}"
        value /= 1024
    return f"{num_bytes} B"


def get_upload_token() -> str:
    token = st.query_params.get("upload_token", "").strip()
    if not token:
        token = uuid.uuid4().hex
        st.query_params["upload_token"] = token
    return token


def get_staging_dir(upload_token: str) -> Path:
    folder = STAGING_ROOT / upload_token
    folder.mkdir(parents=True, exist_ok=True)
    return folder


def load_staging_manifest(upload_token: str) -> dict:
    manifest_path = get_staging_dir(upload_token) / "manifest.json"
    if not manifest_path.exists():
        return {"files": []}
    try:
        return json.loads(manifest_path.read_text(encoding="utf-8"))
    except Exception:
        return {"files": []}


def save_staging_manifest(upload_token: str, manifest: dict) -> None:
    manifest_path = get_staging_dir(upload_token) / "manifest.json"
    manifest_path.write_text(json.dumps(manifest, indent=2), encoding="utf-8")


def clear_staged_files(upload_token: str) -> None:
    folder = get_staging_dir(upload_token)
    if folder.exists():
        shutil.rmtree(folder, ignore_errors=True)


def stage_uploaded_files(upload_token: str, upload_items: list) -> int:
    """Persist uploaded files so they survive Streamlit session reconnects."""
    if not upload_items:
        return 0

    folder = get_staging_dir(upload_token)
    manifest = load_staging_manifest(upload_token)
    existing = {item["fingerprint"] for item in manifest.get("files", [])}
    added = 0

    for item in upload_items:
        if isinstance(item, dict):
            raw_bytes = item["data"]
            name = item["name"]
            content_type = item.get("content_type", "")
        else:
            raw_bytes = bytes(item.getbuffer())
            name = item.name
            content_type = item.type

        fingerprint = hashlib.sha1(raw_bytes).hexdigest()
        if fingerprint in existing:
            continue

        ext = Path(name).suffix
        stored_name = f"{fingerprint}{ext}"
        file_path = folder / stored_name
        file_path.write_bytes(raw_bytes)

        manifest.setdefault("files", []).append(
            {
                "fingerprint": fingerprint,
                "name": name,
                "content_type": content_type,
                "file_path": str(file_path),
                "size": len(raw_bytes),
            }
        )
        existing.add(fingerprint)
        added += 1

    save_staging_manifest(upload_token, manifest)
    return added


def get_staged_files(upload_token: str) -> list[dict]:
    manifest = load_staging_manifest(upload_token)
    staged = []
    changed = False
    for item in manifest.get("files", []):
        path = Path(item.get("file_path", ""))
        if not path.exists():
            changed = True
            continue
        staged.append(item)

    if changed:
        manifest["files"] = staged
        save_staging_manifest(upload_token, manifest)

    return staged


def stage_widget_files(widget_key: str) -> None:
    upload_token = st.session_state.get("upload_token", "")
    if not upload_token:
        return

    upload_items = st.session_state.get(widget_key) or []
    added = stage_uploaded_files(upload_token, upload_items)
    if added > 0:
        st.session_state["upload_stage_notice"] = f"Added {added} file(s) to upload queue."


def get_upload_file_parts(file_item):
    """Return (name, content_type, data, size_bytes) for queue items or Streamlit upload objects."""
    if isinstance(file_item, dict):
        if "file_path" in file_item:
            file_path = Path(file_item["file_path"])
            data = file_path.read_bytes()
            return file_item["name"], file_item.get("content_type", ""), data, int(file_item.get("size", len(data)))
        data = file_item["data"]
        return file_item["name"], file_item.get("content_type", ""), data, len(data)

    data = file_item.getbuffer()
    return file_item.name, file_item.type, data, len(data)


def build_base_output_name(customer_name: str, order_number: str, index: int, original_name: str, content_type: str) -> str:
    safe_customer = sanitize_name(customer_name).replace(" ", "_")
    safe_order = sanitize_name(order_number).replace(" ", "_") if order_number else ""
    ext = guess_extension(original_name, content_type)
    kind = "video" if (content_type or "").startswith("video/") else "photo"
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    if safe_order:
        return f"{safe_customer}_{safe_order}_{timestamp}_{kind}_{index:02d}{ext}"
    return f"{safe_customer}_{timestamp}_{kind}_{index:02d}{ext}"


def ensure_unique_filename(drive_id: str, folder_path: str, candidate_name: str) -> str:
    base = Path(candidate_name).stem
    ext = Path(candidate_name).suffix
    unique_name = candidate_name
    counter = 1

    while True:
        test_path = f"{folder_path.strip('/')}/{unique_name}"
        existing = get_item_by_path(drive_id, test_path)
        if not existing:
            return unique_name
        unique_name = f"{base}_dup{counter:02d}{ext}"
        counter += 1


def upload_small_file(drive_id: str, parent_id: str, output_name: str, data: bytes | memoryview, content_type: str) -> dict:
    encoded_name = quote(output_name, safe="")
    return graph_json(
        "PUT",
        f"/drives/{drive_id}/items/{parent_id}:/{encoded_name}:/content",
        data=data,
        headers={"Content-Type": content_type or "application/octet-stream"},
    )


def upload_large_file(drive_id: str, parent_id: str, output_name: str, data: bytes | memoryview) -> dict:
    encoded_name = quote(output_name, safe="")
    session = graph_json(
        "POST",
        f"/drives/{drive_id}/items/{parent_id}:/{encoded_name}:/createUploadSession",
        json={
            "item": {
                "name": output_name,
                "@microsoft.graph.conflictBehavior": "rename"
            }
        },
    )

    upload_url = session["uploadUrl"]
    file_size = len(data)
    start = 0
    final_response = None

    while start < file_size:
        end = min(start + CHUNK_SIZE, file_size) - 1
        chunk = data[start:end + 1]

        headers = {
            "Content-Length": str(len(chunk)),
            "Content-Range": f"bytes {start}-{end}/{file_size}",
        }

        response = requests.put(upload_url, headers=headers, data=chunk, timeout=300)

        if response.status_code not in (200, 201, 202):
            try:
                detail = response.json()
            except Exception:
                detail = response.text
            raise RuntimeError(f"Large upload failed: {response.status_code} - {detail}")

        final_response = response
        start = end + 1

    if final_response is not None and final_response.content:
        return final_response.json()

    return {"name": output_name}


def send_notification_email(customer_name: str, order_number: str, folder_url: str, uploaded_names: list[str]) -> None:
    if not EMAIL_NOTIFY_ENABLED:
        return
    if not NOTIFY_SENDER_EMAIL or not NOTIFY_TO_EMAIL:
        return

    subject = f"Upload received - {customer_name}"
    if order_number:
        subject += f" - {order_number}"

    uploaded_list_html = "".join([f"<li>{name}</li>" for name in uploaded_names])

    html_body = f"""
    <html>
      <body>
        <p>Hello,</p>
        <p>A new customer upload has been received.</p>
        <p>
          <b>Customer Name:</b> {customer_name}<br>
          <b>Order / Job Number:</b> {order_number or "N/A"}
        </p>
        <p><b>Uploaded files:</b></p>
        <ul>{uploaded_list_html}</ul>
        <p><a href="{folder_url}">Open customer folder in SharePoint</a></p>
      </body>
    </html>
    """

    payload = {
        "message": {
            "subject": subject,
            "body": {
                "contentType": "HTML",
                "content": html_body,
            },
            "toRecipients": [
                {
                    "emailAddress": {
                        "address": NOTIFY_TO_EMAIL
                    }
                }
            ],
        },
        "saveToSentItems": "true",
    }

    graph_request(
        "POST",
        f"/users/{NOTIFY_SENDER_EMAIL}/sendMail",
        json=payload,
        headers={"Content-Type": "application/json"},
    )


# ============================================================
# DIRECT UPLOAD HTML COMPONENT
# Files go browser → SharePoint Graph API, bypassing Streamlit
# WebSocket entirely. This fixes Samsung mobile session drops.
# ============================================================

def build_direct_upload_html(
    access_token: str,
    drive_id: str,
    folder_id: str,
    destination_path: str,
    customer_name: str,
    order_number: str,
    folder_web_url: str,
) -> str:
    # Safely encode values as JSON strings for embedding in JS
    j_token = json.dumps(access_token)
    j_drive = json.dumps(drive_id)
    j_folder = json.dumps(folder_id)
    j_dest = json.dumps(destination_path.strip("/"))
    j_cust = json.dumps(customer_name)
    j_order = json.dumps(order_number)
    j_url = json.dumps(folder_web_url)

    return f"""
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
    background: transparent;
    color: #1a1a2e;
    padding: 0;
  }}
  .card {{
    background: #fff;
    border: 1px solid #e2e8f0;
    border-radius: 14px;
    padding: 18px;
    margin-bottom: 12px;
  }}
  .folder-badge {{
    display: flex;
    align-items: center;
    gap: 8px;
    background: #f0fdf4;
    border: 1px solid #86efac;
    border-radius: 10px;
    padding: 10px 14px;
    margin-bottom: 14px;
    font-size: 14px;
    color: #166534;
    font-weight: 500;
  }}
  .pick-btn {{
    display: block;
    width: 100%;
    padding: 14px;
    background: #0068c9;
    color: #fff;
    font-size: 16px;
    font-weight: 600;
    border: none;
    border-radius: 10px;
    cursor: pointer;
    text-align: center;
    transition: background 0.2s;
    margin-bottom: 10px;
  }}
  .pick-btn:hover {{ background: #0055a3; }}
  .pick-btn:disabled {{ background: #94a3b8; cursor: not-allowed; }}
  .file-list {{
    list-style: none;
    margin: 10px 0;
    max-height: 220px;
    overflow-y: auto;
  }}
  .file-item {{
    display: flex;
    justify-content: space-between;
    align-items: center;
    padding: 8px 10px;
    border-radius: 8px;
    margin-bottom: 4px;
    background: #f8fafc;
    border: 1px solid #e2e8f0;
    font-size: 13px;
  }}
  .file-name {{ font-weight: 500; word-break: break-all; }}
  .file-size {{ color: #64748b; white-space: nowrap; margin-left: 8px; }}
  .file-status {{
    font-size: 12px;
    white-space: nowrap;
    margin-left: 8px;
    font-weight: 600;
  }}
  .status-waiting {{ color: #94a3b8; }}
  .status-uploading {{ color: #0068c9; }}
  .status-done {{ color: #16a34a; }}
  .status-error {{ color: #dc2626; }}
  .progress-bar-wrap {{
    height: 5px;
    background: #e2e8f0;
    border-radius: 3px;
    margin-top: 5px;
    overflow: hidden;
  }}
  .progress-bar-fill {{
    height: 100%;
    background: #0068c9;
    border-radius: 3px;
    width: 0%;
    transition: width 0.3s ease;
  }}
  .upload-btn {{
    display: block;
    width: 100%;
    padding: 15px;
    background: #16a34a;
    color: #fff;
    font-size: 17px;
    font-weight: 700;
    border: none;
    border-radius: 10px;
    cursor: pointer;
    text-align: center;
    margin-top: 12px;
    transition: background 0.2s;
  }}
  .upload-btn:hover {{ background: #15803d; }}
  .upload-btn:disabled {{ background: #94a3b8; cursor: not-allowed; }}
  .result-box {{
    border-radius: 10px;
    padding: 14px;
    margin-top: 12px;
    font-size: 14px;
    display: none;
  }}
  .result-success {{
    background: #f0fdf4;
    border: 1px solid #86efac;
    color: #166534;
  }}
  .result-error {{
    background: #fef2f2;
    border: 1px solid #fca5a5;
    color: #991b1b;
  }}
  .result-title {{
    font-size: 18px;
    font-weight: 700;
    margin-bottom: 6px;
  }}
  .sp-link {{
    display: inline-block;
    margin-top: 10px;
    padding: 9px 16px;
    background: #0068c9;
    color: #fff;
    border-radius: 8px;
    text-decoration: none;
    font-weight: 600;
    font-size: 14px;
  }}
  .overall-progress {{
    margin-top: 12px;
    font-size: 13px;
    color: #475569;
  }}
  .overall-bar-wrap {{
    height: 8px;
    background: #e2e8f0;
    border-radius: 4px;
    margin-top: 6px;
    overflow: hidden;
  }}
  .overall-bar-fill {{
    height: 100%;
    background: linear-gradient(90deg, #0068c9, #38bdf8);
    border-radius: 4px;
    width: 0%;
    transition: width 0.4s ease;
  }}
  .hint {{
    font-size: 12px;
    color: #64748b;
    margin-top: 6px;
  }}
</style>
</head>
<body>

<div class="card">
  <div class="folder-badge">
    ✅ SharePoint folder ready — files will upload directly from your phone.
  </div>

  <input type="file" id="fileInput" multiple
    accept="image/jpeg,image/png,image/heic,image/heif,image/webp,image/dng,
            video/mp4,video/quicktime,video/x-msvideo,video/x-matroska,
            video/3gpp,video/3gpp2,video/mpeg,video/x-ms-wmv,
            .jpg,.jpeg,.png,.heic,.heif,.webp,.dng,
            .mp4,.mov,.m4v,.avi,.mkv,.3gp,.3gpp,.mpeg,.mpg,.wmv"
    style="display:none"
    onchange="onFilesSelected()"
  >

  <button class="pick-btn" id="pickBtn" onclick="document.getElementById('fileInput').click()">
    📷 Select Photos &amp; Videos
  </button>
  <div class="hint">You can select multiple files. On Samsung, tap and hold to multi-select.</div>

  <ul class="file-list" id="fileList"></ul>

  <div class="overall-progress" id="overallProgress" style="display:none">
    <span id="overallLabel">Uploading 0 of 0...</span>
    <div class="overall-bar-wrap"><div class="overall-bar-fill" id="overallBar"></div></div>
  </div>

  <button class="upload-btn" id="uploadBtn" style="display:none" onclick="startUpload()">
    ⬆️ Upload to SharePoint
  </button>

  <div class="result-box result-success" id="successBox">
    <div class="result-title">✅ Upload complete!</div>
    <div id="successDetail"></div>
    <a href="#" id="spLink" class="sp-link" target="_blank">Open folder in SharePoint ↗</a>
  </div>
  <div class="result-box result-error" id="errorBox">
    <div class="result-title">❌ Upload error</div>
    <div id="errorDetail"></div>
  </div>
</div>

<script>
const ACCESS_TOKEN = {j_token};
const DRIVE_ID = {j_drive};
const FOLDER_ID = {j_folder};
const DESTINATION_PATH = {j_dest};  // e.g. "Service/Mike Johnson - 2501878"
const CUSTOMER_NAME = {j_cust};
const ORDER_NUMBER = {j_order};
const FOLDER_URL = {j_url};
const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';
const CHUNK_SIZE = 5 * 1024 * 1024; // 5 MB — safe for mobile

let selectedFiles = [];
let uploadStarted = false;

function fmtSize(bytes) {{
  if (bytes < 1024) return bytes + ' B';
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
  if (bytes < 1024 * 1024 * 1024) return (bytes / 1024 / 1024).toFixed(1) + ' MB';
  return (bytes / 1024 / 1024 / 1024).toFixed(2) + ' GB';
}}

function buildOutputName(file, index) {{
  const now = new Date();
  const pad = n => String(n).padStart(2, '0');
  const date = `${{now.getFullYear()}}${{pad(now.getMonth()+1)}}${{pad(now.getDate())}}`;
  const time = `${{pad(now.getHours())}}${{pad(now.getMinutes())}}${{pad(now.getSeconds())}}`;
  const safeCust = CUSTOMER_NAME.replace(/[^a-zA-Z0-9]/g, '_').replace(/_+/g, '_');
  const orderPart = ORDER_NUMBER ? ORDER_NUMBER.replace(/[^a-zA-Z0-9]/g, '_').replace(/_+/g, '_') + '_' : '';
  const parts = file.name.split('.');
  const ext = parts.length > 1 ? '.' + parts[parts.length - 1].toLowerCase() : '';
  const kind = file.type.startsWith('video/') ? 'video' : 'photo';
  const idx = String(index + 1).padStart(2, '0');
  return `${{safeCust}}_${{orderPart}}${{date}}_${{time}}_${{kind}}_${{idx}}${{ext}}`;
}}

function onFilesSelected() {{
  const input = document.getElementById('fileInput');
  selectedFiles = Array.from(input.files);
  renderFileList();
  document.getElementById('uploadBtn').style.display = selectedFiles.length ? 'block' : 'none';
  document.getElementById('successBox').style.display = 'none';
  document.getElementById('errorBox').style.display = 'none';
}}

function renderFileList() {{
  const list = document.getElementById('fileList');
  list.innerHTML = '';
  selectedFiles.forEach((f, i) => {{
    const li = document.createElement('li');
    li.className = 'file-item';
    li.id = `file-item-${{i}}`;
    li.innerHTML = `
      <div style="flex:1; min-width:0;">
        <div class="file-name">${{f.name}}</div>
        <div class="progress-bar-wrap"><div class="progress-bar-fill" id="bar-${{i}}"></div></div>
      </div>
      <span class="file-size">${{fmtSize(f.size)}}</span>
      <span class="file-status status-waiting" id="status-${{i}}">waiting</span>
    `;
    list.appendChild(li);
  }});
}}

function setFileStatus(i, text, cls) {{
  const el = document.getElementById(`status-${{i}}`);
  if (el) {{
    el.textContent = text;
    el.className = `file-status ${{cls}}`;
  }}
}}

function setFileProgress(i, pct) {{
  const bar = document.getElementById(`bar-${{i}}`);
  if (bar) bar.style.width = (pct * 100).toFixed(1) + '%';
}}

async function createUploadSession(outputName) {{
  const pathParts = DESTINATION_PATH.split('/').map(encodeURIComponent).join('/');
  const encodedName = encodeURIComponent(outputName);
  const url = `${{GRAPH_BASE}}/drives/${{DRIVE_ID}}/root:/${{pathParts}}/${{encodedName}}:/createUploadSession`;
  const resp = await fetch(url, {{
    method: 'POST',
    headers: {{
      'Authorization': `Bearer ${{ACCESS_TOKEN}}`,
      'Content-Type': 'application/json'
    }},
    body: JSON.stringify({{
      item: {{
        '@microsoft.graph.conflictBehavior': 'rename'
      }}
    }})
  }});
  if (!resp.ok) {{
    const txt = await resp.text();
    throw new Error(`Could not start upload session (${{resp.status}}): ${{txt}}`);
  }}
  const data = await resp.json();
  return data.uploadUrl;
}}

async function uploadChunked(uploadUrl, file, onProgress) {{
  const total = file.size;
  let offset = 0;
  while (offset < total) {{
    const end = Math.min(offset + CHUNK_SIZE, total);
    const chunk = file.slice(offset, end);
    const resp = await fetch(uploadUrl, {{
      method: 'PUT',
      headers: {{
        'Content-Length': String(end - offset),
        'Content-Range': `bytes ${{offset}}-${{end - 1}}/${{total}}`
      }},
      body: chunk
    }});
    if (![200, 201, 202].includes(resp.status)) {{
      const txt = await resp.text();
      throw new Error(`Chunk upload failed at byte ${{offset}} (${{resp.status}}): ${{txt}}`);
    }}
    offset = end;
    onProgress(offset / total);
  }}
}}

async function startUpload() {{
  if (uploadStarted) return;
  if (!selectedFiles.length) return;
  uploadStarted = true;

  document.getElementById('uploadBtn').disabled = true;
  document.getElementById('pickBtn').disabled = true;
  document.getElementById('uploadBtn').textContent = '⏳ Uploading...';
  document.getElementById('overallProgress').style.display = 'block';
  document.getElementById('successBox').style.display = 'none';
  document.getElementById('errorBox').style.display = 'none';

  const total = selectedFiles.length;
  let done = 0;
  const errors = [];

  for (let i = 0; i < selectedFiles.length; i++) {{
    const file = selectedFiles[i];
    const outputName = buildOutputName(file, i);
    document.getElementById('overallLabel').textContent = `Uploading ${{i + 1}} of ${{total}}: ${{file.name}}`;
    document.getElementById('overallBar').style.width = `${{(i / total * 100).toFixed(1)}}%`;
    setFileStatus(i, 'uploading...', 'status-uploading');
    setFileProgress(i, 0);

    try {{
      const uploadUrl = await createUploadSession(outputName);
      await uploadChunked(uploadUrl, file, pct => setFileProgress(i, pct));
      setFileStatus(i, '✓ done', 'status-done');
      setFileProgress(i, 1);
      done++;
    }} catch (err) {{
      setFileStatus(i, '✗ failed', 'status-error');
      errors.push(`${{file.name}}: ${{err.message}}`);
      console.error('Upload error', err);
    }}
  }}

  document.getElementById('overallBar').style.width = '100%';
  document.getElementById('overallLabel').textContent =
    errors.length === 0
      ? `✅ All ${{total}} file(s) uploaded!`
      : `${{done}} of ${{total}} uploaded, ${{errors.length}} failed.`;

  document.getElementById('uploadBtn').textContent = '⬆️ Upload to SharePoint';
  document.getElementById('uploadBtn').disabled = false;
  document.getElementById('pickBtn').disabled = false;
  uploadStarted = false;

  if (done > 0) {{
    const successBox = document.getElementById('successBox');
    successBox.style.display = 'block';
    document.getElementById('successDetail').textContent =
      `${{done}} file(s) saved to SharePoint successfully.${{errors.length ? ' Some files had errors — see list above.' : ''}}`;
    const link = document.getElementById('spLink');
    if (FOLDER_URL) {{
      link.href = FOLDER_URL;
      link.style.display = 'inline-block';
    }} else {{
      link.style.display = 'none';
    }}
  }}

  if (errors.length > 0) {{
    const errBox = document.getElementById('errorBox');
    errBox.style.display = 'block';
    document.getElementById('errorDetail').innerHTML =
      '<ul style="margin-top:6px;padding-left:18px">' +
      errors.map(e => `<li style="margin-bottom:4px">${{e}}</li>`).join('') +
      '</ul>';
  }}
}}
</script>
</body>
</html>
"""


# ============================================================
# MAIN FORM
# ============================================================

validate_config()

upload_token = get_upload_token()
st.session_state["upload_token"] = upload_token

customer_name = st.text_input("Customer name *", placeholder="Example: Mike Johnson")
order_number = st.text_input("Order / job number", placeholder="Example: 2501878")
notes = st.text_area("Notes (optional)", placeholder="Optional notes for your own reference")

upload_mode = st.radio(
    "Upload mode",
    [
        "📱 Samsung / Mobile Direct Upload  ← Use this on phones",
        "Mobile multi-select (photos + videos separately)",
        "Standard multi-select",
        "Phone-safe (add one by one)",
    ],
    horizontal=False,
    help="Samsung Direct Upload sends files straight from your browser to SharePoint — no Streamlit connection needed. Fixes all Samsung session drop issues.",
)

allowed_types = [
    "jpg", "jpeg", "png", "heic", "heif", "webp", "dng",
    "mp4", "mov", "m4v", "avi", "mkv", "3gp", "3gpp", "mpeg", "mpg", "wmv",
]

# ============================================================
# MODE: Samsung / Mobile Direct Upload
# Files go: Browser → Graph API → SharePoint (no WebSocket)
# ============================================================
if upload_mode == "📱 Samsung / Mobile Direct Upload  ← Use this on phones":

    ready = st.session_state.get("direct_upload_ready")

    if not ready:
        st.info(
            "**How it works:** Fill in the customer name above, then tap **Prepare Upload Folder**. "
            "The app creates the SharePoint folder and opens a file picker. "
            "Your phone uploads files directly to SharePoint — no connection drops.",
            icon="ℹ️",
        )

        if st.button("Prepare Upload Folder", use_container_width=True, type="primary"):
            if not customer_name.strip():
                st.error("Please enter the customer name first.")
            else:
                with st.spinner("Connecting to SharePoint and creating folder..."):
                    try:
                        site_id = get_site_id()
                        drive_id = get_drive_id(site_id)

                        folder_name = sanitize_name(customer_name)
                        if order_number.strip():
                            folder_name = f"{folder_name} - {sanitize_name(order_number)}"

                        base_folder = BASE_FOLDER_NAME.strip().strip("/")
                        destination_path = f"{base_folder}/{folder_name}" if base_folder else folder_name

                        folder_item = ensure_folder_path(drive_id, destination_path)
                        folder_web_url = folder_item.get("webUrl", "")
                        token = get_access_token()

                        # Upload notes as a text file if provided
                        if notes.strip():
                            try:
                                notes_filename = f"{sanitize_name(customer_name).replace(' ', '_')}_notes.txt"
                                upload_small_file(
                                    drive_id=drive_id,
                                    parent_id=folder_item["id"],
                                    output_name=notes_filename,
                                    data=notes.strip().encode("utf-8"),
                                    content_type="text/plain",
                                )
                            except Exception as ne:
                                st.warning(f"Notes could not be saved: {ne}")

                        st.session_state["direct_upload_ready"] = {
                            "token": token,
                            "drive_id": drive_id,
                            "folder_id": folder_item["id"],
                            "destination_path": destination_path,
                            "folder_web_url": folder_web_url,
                            "folder_name": folder_name,
                            "customer_name": customer_name,
                            "order_number": order_number,
                        }
                        st.rerun()

                    except Exception as exc:
                        st.error(f"Could not prepare SharePoint folder: {exc}")

    else:
        st.success(
            f"📁 Folder ready: **{ready['folder_name']}**  \n"
            "Select your photos and videos below, then tap Upload to SharePoint."
        )

        components.html(
            build_direct_upload_html(
                access_token=ready["token"],
                drive_id=ready["drive_id"],
                folder_id=ready["folder_id"],
                destination_path=ready["destination_path"],
                customer_name=ready["customer_name"],
                order_number=ready["order_number"],
                folder_web_url=ready["folder_web_url"],
            ),
            height=620,
            scrolling=True,
        )

        col1, col2 = st.columns([1, 1])
        with col1:
            if ready.get("folder_web_url"):
                st.link_button("Open folder in SharePoint ↗", ready["folder_web_url"], use_container_width=True)
        with col2:
            if st.button("Start over / New customer", use_container_width=True):
                del st.session_state["direct_upload_ready"]
                st.rerun()

    st.stop()  # Skip the regular upload button for this mode


# ============================================================
# MODE: Phone-safe (one at a time)
# ============================================================
elif upload_mode == "Phone-safe (add one by one)":
    st.caption("For Samsung/mobile issues: pick one file, tap Add file, repeat, then Upload.")
    single_file = st.file_uploader(
        "Choose one photo or video",
        type=allowed_types,
        accept_multiple_files=False,
        key="single_file_input",
        help="Add one item at a time to the queue. Supports common Samsung formats (HEIC/HEIF/DNG/MP4/MKV/3GP).",
    )

    add_col, clear_col = st.columns([1, 1])
    with add_col:
        add_clicked = st.button("Add selected file", use_container_width=True)
    with clear_col:
        clear_clicked = st.button("Clear queued files", use_container_width=True)

    if clear_clicked:
        clear_staged_files(upload_token)
        st.rerun()

    if add_clicked:
        if not single_file:
            st.warning("Select a file first, then tap Add selected file.")
        else:
            item = {
                "name": single_file.name,
                "content_type": single_file.type,
                "data": bytes(single_file.getbuffer()),
            }
            added = stage_uploaded_files(upload_token, [item])
            if added > 0:
                st.info("Added 1 file to upload queue.")

    files_to_upload = get_staged_files(upload_token)
    selected_count = len(files_to_upload)
    st.caption(f"Queued files: {selected_count}")

    if files_to_upload:
        with st.expander("Queued files", expanded=False):
            for idx, f in enumerate(files_to_upload, start=1):
                st.write(f"{idx}. {f['name']} ({format_size(int(f.get('size', 0)))})")

# ============================================================
# MODE: Mobile multi-select
# ============================================================
elif upload_mode == "Mobile multi-select (photos + videos separately)":
    st.caption("Samsung-friendly mode: select multiple photos first, then multiple videos, then upload once.")

    st.file_uploader(
        "Select photos (multi-select)",
        type=["jpg", "jpeg", "png", "heic", "heif", "webp", "dng"],
        accept_multiple_files=True,
        key="mobile_photo_multi",
        help="Select one or more photos in a single action.",
        on_change=stage_widget_files,
        args=("mobile_photo_multi",),
    )

    st.file_uploader(
        "Select videos (multi-select)",
        type=["mp4", "mov", "m4v", "avi", "mkv", "3gp", "3gpp", "mpeg", "mpg", "wmv"],
        accept_multiple_files=True,
        key="mobile_video_multi",
        help="Select one or more videos in a single action.",
        on_change=stage_widget_files,
        args=("mobile_video_multi",),
    )

    files_to_upload = get_staged_files(upload_token)
    selected_count = len(files_to_upload)
    st.caption(f"Selected files: {selected_count}")

    if files_to_upload:
        with st.expander("Queued files", expanded=False):
            for idx, f in enumerate(files_to_upload, start=1):
                st.write(f"{idx}. {f['name']} ({format_size(int(f.get('size', 0)))})")

    if st.button("Clear selected files", use_container_width=True, key="mobile_multi_clear"):
        clear_staged_files(upload_token)
        st.rerun()

# ============================================================
# MODE: Standard multi-select
# ============================================================
else:
    st.file_uploader(
        "Upload photos and videos *",
        type=allowed_types,
        accept_multiple_files=True,
        help=(
            f"You can select multiple files at once. Max {MAX_PER_FILE_UPLOAD_MB} MB per file. "
            f"For best phone reliability: up to {MAX_FILES_PER_BATCH} files per batch and around {RECOMMENDED_MOBILE_TOTAL_MB} MB total."
        ),
        key="standard_multi_files",
        on_change=stage_widget_files,
        args=("standard_multi_files",),
    )

    files_to_upload = get_staged_files(upload_token)
    selected_count = len(files_to_upload)
    st.caption(f"Selected files: {selected_count}")

    if files_to_upload:
        with st.expander("Queued files", expanded=False):
            for idx, f in enumerate(files_to_upload, start=1):
                st.write(f"{idx}. {f['name']} ({format_size(int(f.get('size', 0)))})")

    if st.button("Clear selected files", use_container_width=True, key="standard_multi_clear"):
        clear_staged_files(upload_token)
        st.rerun()


notice = st.session_state.pop("upload_stage_notice", "")
if notice:
    st.info(notice)

# ============================================================
# REGULAR UPLOAD BUTTON (non-Samsung modes)
# ============================================================
submitted = st.button("Upload to SharePoint", use_container_width=True)

if submitted:
    customer_name = customer_name.strip()
    order_number = order_number.strip()
    notes = notes.strip()

    if not customer_name:
        st.error("Please enter the customer name.")
        st.stop()

    if not files_to_upload:
        st.error("Please upload at least one file.")
        st.stop()

    if len(files_to_upload) > MAX_FILES_PER_BATCH:
        st.error(
            f"You selected {len(files_to_upload)} files. "
            f"Please upload up to {MAX_FILES_PER_BATCH} files at a time for reliable phone uploads."
        )
        st.stop()

    max_per_file_bytes = MAX_PER_FILE_UPLOAD_MB * 1024 * 1024
    oversized_files = []
    for file_item in files_to_upload:
        file_name, _, _, file_size = get_upload_file_parts(file_item)
        if file_size > max_per_file_bytes:
            oversized_files.append((file_name, file_size))

    if oversized_files:
        st.error(
            f"One or more files are larger than the allowed {MAX_PER_FILE_UPLOAD_MB} MB per file. "
            "Remove large files and try again."
        )
        for file_name, file_size in oversized_files[:5]:
            st.write(f"- {file_name}: {format_size(file_size)}")
        if len(oversized_files) > 5:
            st.write(f"- ...and {len(oversized_files) - 5} more file(s)")
        st.stop()

    total_size_bytes = sum(get_upload_file_parts(file_item)[3] for file_item in files_to_upload)
    max_total_bytes = MAX_TOTAL_UPLOAD_MB * 1024 * 1024
    if total_size_bytes > max_total_bytes:
        st.error(
            f"Total selected size is {format_size(total_size_bytes)}. "
            f"Please keep total selection under {MAX_TOTAL_UPLOAD_MB} MB and upload in batches."
        )
        st.stop()

    recommended_total_bytes = RECOMMENDED_MOBILE_TOTAL_MB * 1024 * 1024
    if total_size_bytes > recommended_total_bytes:
        st.warning(
            f"You selected {format_size(total_size_bytes)} total. "
            f"Phone uploads are more reliable around {RECOMMENDED_MOBILE_TOTAL_MB} MB or less per batch."
        )

    try:
        site_id = get_site_id()
        drive_id = get_drive_id(site_id)

        folder_name = sanitize_name(customer_name)
        if order_number:
            folder_name = f"{folder_name} - {sanitize_name(order_number)}"

        base_folder = BASE_FOLDER_NAME.strip().strip("/")
        destination_path = f"{base_folder}/{folder_name}" if base_folder else folder_name

        with st.spinner("Creating folder and uploading files..."):
            destination_folder = ensure_folder_path(drive_id, destination_path)

            if notes:
                try:
                    notes_filename = ensure_unique_filename(
                        drive_id,
                        destination_path,
                        f"{sanitize_name(customer_name).replace(' ', '_')}_notes.txt"
                    )
                    notes_content = notes.encode("utf-8")
                    upload_small_file(
                        drive_id=drive_id,
                        parent_id=destination_folder["id"],
                        output_name=notes_filename,
                        data=notes_content,
                        content_type="text/plain",
                    )
                except Exception as notes_exc:
                    st.warning(f"Notes could not be uploaded. Continuing with photos/videos. Details: {notes_exc}")

            progress = st.progress(0)
            status = st.empty()
            uploaded_names = []

            for i, file_item in enumerate(files_to_upload, start=1):
                original_name, content_type, data, _ = get_upload_file_parts(file_item)
                candidate_name = build_base_output_name(
                    customer_name=customer_name,
                    order_number=order_number,
                    index=i,
                    original_name=original_name,
                    content_type=content_type,
                )
                output_name = ensure_unique_filename(drive_id, destination_path, candidate_name)

                status.write(f"Uploading {i}/{len(files_to_upload)}: {output_name}")

                if len(data) <= SMALL_UPLOAD_LIMIT:
                    upload_small_file(
                        drive_id=drive_id,
                        parent_id=destination_folder["id"],
                        output_name=output_name,
                        data=data,
                        content_type=content_type,
                    )
                else:
                    upload_large_file(
                        drive_id=drive_id,
                        parent_id=destination_folder["id"],
                        output_name=output_name,
                        data=data,
                    )

                uploaded_names.append(output_name)
                progress.progress(i / len(files_to_upload))

        final_folder = get_item_by_path(drive_id, destination_path)
        folder_url = final_folder.get("webUrl") if final_folder else ""

        try:
            send_notification_email(
                customer_name=customer_name,
                order_number=order_number,
                folder_url=folder_url,
                uploaded_names=uploaded_names,
            )
        except Exception as notify_exc:
            st.warning(f"Files uploaded, but email notification failed: {notify_exc}")

        st.success(f"Done. Uploaded {len(uploaded_names)} file(s) successfully.")

        col1, col2 = st.columns([1, 1])
        with col1:
            st.metric("Files uploaded", len(uploaded_names))
        with col2:
            st.metric("Customer folder", folder_name)

        if folder_url:
            st.link_button("Open customer folder in SharePoint", folder_url, use_container_width=True)

        with st.expander("Uploaded files", expanded=True):
            for name in uploaded_names:
                st.write(f"- {name}")

        st.markdown(
            '<div class="footer-note">Tip: You can bookmark this page on the phone home screen for faster access.</div>',
            unsafe_allow_html=True,
        )

        if upload_mode == "Phone-safe (add one by one)":
            st.session_state.mobile_upload_queue = []
        clear_staged_files(upload_token)

    except Exception as exc:
        st.error(str(exc))
