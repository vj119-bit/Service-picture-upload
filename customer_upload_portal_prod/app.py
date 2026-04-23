import mimetypes
import os
import re
from datetime import datetime
from pathlib import Path
from urllib.parse import quote

import requests
import streamlit as st
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
MAX_TOTAL_UPLOAD_MB = 5000


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
# MAIN FORM
# ============================================================

validate_config()

with st.form("upload_form", clear_on_submit=False):
    customer_name = st.text_input("Customer name *", placeholder="Example: Mike Johnson")
    order_number = st.text_input("Order / job number", placeholder="Example: 2501878")
    notes = st.text_area("Notes (optional)", placeholder="Optional notes for your own reference")
    uploaded_files = st.file_uploader(
        "Upload photos and videos *",
        type=["jpg", "jpeg", "png", "heic", "heif", "webp", "mp4", "mov", "m4v", "avi"],
        accept_multiple_files=True,
        help="You can select multiple files at once. Max 500 MB per file.",
    )
    submitted = st.form_submit_button("Upload to SharePoint", use_container_width=True)


if submitted:
    customer_name = customer_name.strip()
    order_number = order_number.strip()
    notes = notes.strip()

    if not customer_name:
        st.error("Please enter the customer name.")
        st.stop()

    if not uploaded_files:
        st.error("Please upload at least one file.")
        st.stop()

    total_size_bytes = sum(getattr(file, "size", 0) or 0 for file in uploaded_files)
    max_total_bytes = MAX_TOTAL_UPLOAD_MB * 1024 * 1024
    if total_size_bytes > max_total_bytes:
        st.error(
            f"Total selected size is {format_size(total_size_bytes)}. "
            f"Please keep total selection under {MAX_TOTAL_UPLOAD_MB} MB and upload in batches."
        )
        st.stop()

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

            progress = st.progress(0)
            status = st.empty()
            uploaded_names = []

            for i, file in enumerate(uploaded_files, start=1):
                data = file.getbuffer()
                candidate_name = build_base_output_name(
                    customer_name=customer_name,
                    order_number=order_number,
                    index=i,
                    original_name=file.name,
                    content_type=file.type,
                )
                output_name = ensure_unique_filename(drive_id, destination_path, candidate_name)

                status.write(f"Uploading {i}/{len(uploaded_files)}: {output_name}")

                if len(data) <= SMALL_UPLOAD_LIMIT:
                    upload_small_file(
                        drive_id=drive_id,
                        parent_id=destination_folder["id"],
                        output_name=output_name,
                        data=data,
                        content_type=file.type,
                    )
                else:
                    upload_large_file(
                        drive_id=drive_id,
                        parent_id=destination_folder["id"],
                        output_name=output_name,
                        data=data,
                    )

                uploaded_names.append(output_name)
                progress.progress(i / len(uploaded_files))

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

    except Exception as exc:
        st.error(str(exc))
