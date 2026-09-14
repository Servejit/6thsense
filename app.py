# ============================================================
# 6THSENSE (6S-FO200) VARDAAN
# STREAMLIT + SUPABASE
#
# MASTER EXCEL -> VALUE-ONLY OUTPUT + CONDITIONS
#
# ADMIN:
#   - Upload Master
#   - Replace Master
#   - Delete Master
#   - Generate / Download
#   - User Management
#   - Telegram Settings
#
# USER:
#   - View
#   - Generate
#   - Download
#
# SECURITY:
#   - Supabase Authentication
#   - Supabase Database
#   - Supabase Private Storage
#   - RLS
#   - Server-side secret key only
#   - No public signup
# ============================================================

import io
import re
import os
import json
import warnings
from datetime import datetime

import pandas as pd
import numpy as np
import requests
import streamlit as st

from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

from supabase import create_client, Client
from supabase.lib.client_options import ClientOptions


warnings.filterwarnings("ignore")


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="6thSense (6S-FO200) Vardaan",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)


# ============================================================
# CONSTANTS
# ============================================================

APP_NAME = "6thSense (6S-FO200) Vardaan"

MASTER_BUCKET = "master-files"
MASTER_PATH = "master/master.xlsx"

OUTPUT_NAME = "6thsense(6S-FO200)Vardaan.xlsx"

ALLOWED_EXTENSIONS = [
    ".xlsx",
    ".xlsm",
    ".xltx",
    ".xltm"
]


# ============================================================
# SUPABASE CLIENTS
# ============================================================

def get_supabase_url():

    return st.secrets["SUPABASE_URL"]


def get_publishable_key():

    if "SUPABASE_PUBLISHABLE_KEY" in st.secrets:
        return st.secrets["SUPABASE_PUBLISHABLE_KEY"]

    # Backward-compatible name
    if "SUPABASE_ANON_KEY" in st.secrets:
        return st.secrets["SUPABASE_ANON_KEY"]

    raise RuntimeError(
        "SUPABASE_PUBLISHABLE_KEY is missing from Streamlit Secrets."
    )


def get_service_key():

    if "SUPABASE_SECRET_KEY" in st.secrets:
        return st.secrets["SUPABASE_SECRET_KEY"]

    # Backward-compatible legacy name
    if "SUPABASE_SERVICE_ROLE_KEY" in st.secrets:
        return st.secrets["SUPABASE_SERVICE_ROLE_KEY"]

    raise RuntimeError(
        "SUPABASE_SECRET_KEY is missing from Streamlit Secrets."
    )


@st.cache_resource
def get_public_client() -> Client:

    return create_client(
        get_supabase_url(),
        get_publishable_key()
    )


@st.cache_resource
def get_admin_client() -> Client:

    return create_client(
        get_supabase_url(),
        get_service_key(),
        options=ClientOptions(
            auto_refresh_token=False,
            persist_session=False
        )
    )


supabase = get_public_client()


# ============================================================
# SESSION STATE
# ============================================================

DEFAULT_SESSION = {
    "access_token": None,
    "refresh_token": None,
    "user": None,
    "role": None,
    "login_email": None
}


for key, value in DEFAULT_SESSION.items():

    if key not in st.session_state:

        st.session_state[key] = value


# ============================================================
# BASIC HELPERS
# ============================================================

def safe_text(value):

    if value is None:
        return ""

    return str(value).strip()


def normalize_header(value):

    if value is None:
        return ""

    value = str(value).strip().lower()

    value = re.sub(
        r"\s+",
        " ",
        value
    )

    return value


def numeric_value(value):

    if value is None:
        return None

    if isinstance(
        value,
        (int, float, np.number)
    ):

        if pd.isna(value):
            return None

        return float(value)

    text = str(value).strip()

    if not text:
        return None

    text = text.replace(",", "")
    text = text.replace("%", "")

    text = re.sub(
        r"[^0-9eE+\-.]",
        "",
        text
    )

    if not text:
        return None

    try:

        return float(text)

    except Exception:

        return None


def number_inside_parentheses(value):

    if value is None:
        return None

    text = str(value)

    match = re.search(
        r"\(\s*([-+]?\d+(?:\.\d+)?)\s*\)",
        text
    )

    if not match:
        return None

    try:

        return float(
            match.group(1)
        )

    except Exception:

        return None


def first_number(value):

    if value is None:
        return None

    text = str(value)

    text_without_parentheses = re.sub(
        r"\([^)]*\)",
        "",
        text
    )

    match = re.search(
        r"[-+]?\d+(?:\.\d+)?",
        text_without_parentheses
    )

    if not match:
        return None

    try:

        return float(
            match.group(0)
        )

    except Exception:

        return None


# ============================================================
# DATE DETECTION
# ============================================================

MONTHS = {
    "jan": 1,
    "feb": 2,
    "mar": 3,
    "apr": 4,
    "may": 5,
    "jun": 6,
    "jul": 7,
    "aug": 8,
    "sep": 9,
    "oct": 10,
    "nov": 11,
    "dec": 12
}


def extract_date_from_heading(value):

    if value is None:
        return None

    text = str(value).strip().lower()

    match = re.search(
        r"\b(\d{1,2})\s*[-/ ]\s*"
        r"(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)"
        r"[a-z]*\b",
        text
    )

    if match:

        day = int(
            match.group(1)
        )

        month = MONTHS[
            match.group(2)[:3]
        ]

        try:

            return datetime(
                datetime.now().year,
                month,
                day
            )

        except Exception:

            return None


    match = re.search(
        r"\b(\d{1,2})\s*/\s*(\d{1,2})\b",
        text
    )

    if match:

        day = int(
            match.group(1)
        )

        month = int(
            match.group(2)
        )

        try:

            return datetime(
                datetime.now().year,
                month,
                day
            )

        except Exception:

            return None

    return None


def find_recent_p2l_column(ws):

    candidates = []

    for cell in ws[1]:

        if cell.value is None:
            continue

        heading = str(
            cell.value
        ).strip().lower()

        if not (
            re.search(
                r"\bp2l\b",
                heading
            )
            or
            re.search(
                r"\bo2l\b",
                heading
            )
        ):

            continue

        dt = extract_date_from_heading(
            heading
        )

        if dt is not None:

            candidates.append(
                (
                    dt,
                    cell.column,
                    str(cell.value)
                )
            )

    if not candidates:
        return None

    candidates.sort(
        key=lambda x: x[0],
        reverse=True
    )

    return candidates[0][1]


# ============================================================
# EXCEL HELPERS
# ============================================================

def find_heading(ws, heading):

    target = normalize_header(
        heading
    )

    for cell in ws[1]:

        if normalize_header(
            cell.value
        ) == target:

            return cell.column

    return None


def get_fill(hex_color):

    return PatternFill(
        fill_type="solid",
        fgColor=str(
            hex_color
        ).replace(
            "#",
            ""
        ).upper()
    )


def set_row_green(
    ws,
    row_number,
    green_fill
):

    for col in range(
        1,
        ws.max_column + 1
    ):

        ws.cell(
            row=row_number,
            column=col
        ).fill = green_fill


# ============================================================
# GET SETTINGS
# ============================================================

def get_settings():

    try:

        response = (
            get_admin_client()
            .table("app_settings")
            .select("*")
            .eq("id", 1)
            .limit(1)
            .execute()
        )

        if response.data:

            return response.data[0]

    except Exception:

        pass

    return {
        "id": 1,
        "light_blue": "ADD8E6",
        "light_green": "90EE90",
        "telegram_enabled": False,
        "telegram_bot_token": "",
        "telegram_chat_id": "",
        "alert_enabled": False,
        "alert_type": "O2L",
        "alert_operator": "<",
        "alert_value": -1.00
    }


# ============================================================
# AUTHENTICATION
# ============================================================

def get_authenticated_user():

    access_token = st.session_state.get(
        "access_token"
    )

    if not access_token:
        return None

    try:

        response = (
            supabase.auth.get_user(
                access_token
            )
        )

        return response.user

    except Exception:

        return None


def get_user_role(user_id):

    try:

        response = (
            get_admin_client()
            .table("profiles")
            .select("role, email, is_active")
            .eq("id", user_id)
            .limit(1)
            .execute()
        )

        if not response.data:
            return None

        profile = response.data[0]

        if not profile.get(
            "is_active",
            True
        ):

            return None

        return profile.get(
            "role",
            "user"
        )

    except Exception:

        return None


def login_user(
    email,
    password
):

    try:

        response = (
            supabase.auth
            .sign_in_with_password(
                {
                    "email": email.strip(),
                    "password": password
                }
            )
        )

        session = response.session
        user = response.user

        if session is None or user is None:

            return False, (
                "Login failed."
            )

        # Store only session tokens in Streamlit
        st.session_state.access_token = (
            session.access_token
        )

        st.session_state.refresh_token = (
            session.refresh_token
        )

        st.session_state.user = user
        st.session_state.login_email = (
            user.email
        )

        role = get_user_role(
            user.id
        )

        if role is None:

            # Remove session if no valid
            # profile / disabled account
            try:
                supabase.auth.sign_out()
            except Exception:
                pass

            st.session_state.access_token = None
            st.session_state.refresh_token = None
            st.session_state.user = None
            st.session_state.role = None

            return False, (
                "Your account is not authorized "
                "for this application."
            )

        st.session_state.role = role

        return True, "Login successful."

    except Exception as e:

        return False, str(e)


def logout_user():

    try:

        if (
            st.session_state.get(
                "access_token"
            )
        ):

            supabase.auth.sign_out()

    except Exception:

        pass

    for key in [
        "access_token",
        "refresh_token",
        "user",
        "role",
        "login_email"
    ]:

        st.session_state[key] = None

    st.rerun()


# ============================================================
# REFRESH SESSION
# ============================================================

def refresh_session():

    access_token = st.session_state.get(
        "access_token"
    )

    refresh_token = st.session_state.get(
        "refresh_token"
    )

    if not access_token:
        return False

    try:

        if refresh_token:

            response = (
                supabase.auth.set_session(
                    access_token,
                    refresh_token
                )
            )

            if response.session:

                st.session_state.access_token = (
                    response.session.access_token
                )

                st.session_state.refresh_token = (
                    response.session.refresh_token
                )

        user = get_authenticated_user()

        if user is None:
            return False

        role = get_user_role(
            user.id
        )

        if role is None:
            return False

        st.session_state.user = user
        st.session_state.role = role
        st.session_state.login_email = user.email

        return True

    except Exception:

        return False


# ============================================================
# ADMIN CHECK
# ============================================================

def is_admin():

    return (
        st.session_state.get(
            "role"
        ) == "admin"
    )


def require_admin():

    if not is_admin():

        st.error(
            "Administrator access required."
        )

        st.stop()


# ============================================================
# MASTER FILE METADATA
# ============================================================

def get_master_metadata():

    try:

        response = (
            get_admin_client()
            .table("master_files")
            .select("*")
            .eq("id", 1)
            .limit(1)
            .execute()
        )

        if response.data:

            return response.data[0]

    except Exception:

        pass

    return None


# ============================================================
# DOWNLOAD MASTER FROM PRIVATE STORAGE
# ============================================================

def download_master_bytes():

    metadata = get_master_metadata()

    if not metadata:

        return None

    if not metadata.get(
        "storage_path"
    ):

        return None

    try:

        data = (
            get_admin_client()
            .storage
            .from_(MASTER_BUCKET)
            .download(
                metadata["storage_path"]
            )
        )

        return bytes(data)

    except Exception as e:

        st.error(
            f"Unable to read Master file: {e}"
        )

        return None


# ============================================================
# UPLOAD / REPLACE MASTER
# ============================================================

def save_master_file(
    uploaded_file,
    current_user_id
):

    require_admin()

    filename = uploaded_file.name

    extension = os.path.splitext(
        filename
    )[1].lower()

    if extension not in ALLOWED_EXTENSIONS:

        return False, (
            "Invalid file type. "
            "Please upload an Excel file."
        )

    file_bytes = uploaded_file.getvalue()

    if not file_bytes:

        return False, (
            "The uploaded file is empty."
        )

    admin = get_admin_client()

    try:

        # Remove old file first if it exists
        try:

            admin.storage \
                .from_(MASTER_BUCKET) \
                .remove([
                    MASTER_PATH
                ])

        except Exception:

            pass

        # Upload replacement
        admin.storage \
            .from_(MASTER_BUCKET) \
            .upload(
                MASTER_PATH,
                file_bytes,
                {
                    "content-type":
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    "cache-control":
                        "0",
                    "upsert":
                        "true"
                }
            )

        # Validate workbook before committing metadata
        test_wb = load_workbook(
            io.BytesIO(file_bytes),
            data_only=True,
            read_only=True
        )

        if "summary" not in test_wb.sheetnames:

            test_wb.close()

            # Remove invalid file
            try:

                admin.storage \
                    .from_(MASTER_BUCKET) \
                    .remove([
                        MASTER_PATH
                    ])

            except Exception:
                pass

            return False, (
                "Master file must contain "
                "a sheet named 'summary'."
            )

        test_wb.close()

        # Calculate SHA-256 for version tracking
        import hashlib

        file_hash = hashlib.sha256(
            file_bytes
        ).hexdigest()

        metadata = {
            "id": 1,
            "storage_path": MASTER_PATH,
            "original_filename": filename,
            "file_size": len(file_bytes),
            "sha256": file_hash,
            "uploaded_by": current_user_id,
            "uploaded_at": datetime.utcnow().isoformat()
        }

        admin.table(
            "master_files"
        ).upsert(
            metadata
        ).execute()

        return True, (
            "Master file uploaded/replaced successfully."
        )

    except Exception as e:

        return False, str(e)


# ============================================================
# DELETE MASTER
# ============================================================

def delete_master():

    require_admin()

    admin = get_admin_client()

    try:

        try:

            admin.storage \
                .from_(MASTER_BUCKET) \
                .remove([
                    MASTER_PATH
                ])

        except Exception:
            pass

        admin.table(
            "master_files"
        ).delete().eq(
            "id",
            1
        ).execute()

        return True, (
            "Master file deleted successfully."
        )

    except Exception as e:

        return False, str(e)


# ============================================================
# CREATE VALUE-ONLY OUTPUT
#
# THIS PRESERVES THE ORIGINAL COLAB LOGIC.
# ============================================================

def create_output(
    master_bytes
):

    if not master_bytes:

        raise ValueError(
            "No Master file available."
        )

    source_wb = load_workbook(
        io.BytesIO(master_bytes),
        data_only=True
    )

    if "summary" not in source_wb.sheetnames:

        source_wb.close()

        raise ValueError(
            "The Master file does not contain "
            "a sheet named 'summary'."
        )

    source_ws = source_wb["summary"]

    output_wb = Workbook()

    output_ws = output_wb.active
    output_ws.title = "summary"

    # ========================================================
    # COPY ONLY VALUES
    # ========================================================

    for row in source_ws.iter_rows():

        for source_cell in row:

            output_ws.cell(
                row=source_cell.row,
                column=source_cell.column,
                value=source_cell.value
            )

    output_ws.freeze_panes = "A2"

    output_ws.auto_filter.ref = (
        f"A1:{get_column_letter(output_ws.max_column)}"
        f"{output_ws.max_row}"
    )

    for cell in output_ws[1]:

        cell.font = Font(
            bold=True
        )

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

    # ========================================================
