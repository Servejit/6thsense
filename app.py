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
    # FIND COLUMNS
    # ========================================================

    col_sum_i = find_heading(
        output_ws,
        "Sum I"
    )

    col_16gt = find_heading(
        output_ws,
        "16> C-B / Avg.4"
    )

    col_16lt = find_heading(
        output_ws,
        "16< D-B / Avg.4"
    )

    col_o2h = find_heading(
        output_ws,
        "Sum O2H.10"
    )

    col_o2l = find_heading(
        output_ws,
        "Sum O2L.10"
    )

    col_recent_p2l = find_recent_p2l_column(
        output_ws
    )

    # ========================================================
    # COLORS
    # ========================================================

    settings = get_settings()

    light_blue = settings.get(
        "light_blue",
        "ADD8E6"
    )

    light_green = settings.get(
        "light_green",
        "90EE90"
    )

    blue_fill = get_fill(
        light_blue
    )

    green_fill = get_fill(
        light_green
    )

    condition_count = {}

    for row in range(
        2,
        output_ws.max_row + 1
    ):

        condition_count[row] = 0

    # ========================================================
    # CONDITION 1
    #
    # Sum I < -4.00
    # ========================================================

    if col_sum_i:

        for row in range(
            2,
            output_ws.max_row + 1
        ):

            value = numeric_value(
                output_ws.cell(
                    row=row,
                    column=col_sum_i
                ).value
            )

            if (
                value is not None
                and value < -4.00
            ):

                output_ws.cell(
                    row=row,
                    column=col_sum_i
                ).fill = blue_fill

                output_ws.cell(
                    row=row,
                    column=1
                ).fill = blue_fill

                condition_count[row] += 1

    # ========================================================
    # CONDITION 2
    #
    # 16> C-B / Avg.4
    #
    # Parentheses < 0.50
    # ========================================================

    if col_16gt:

        for row in range(
            2,
            output_ws.max_row + 1
        ):

            cell = output_ws.cell(
                row=row,
                column=col_16gt
            )

            avg_value = (
                number_inside_parentheses(
                    cell.value
                )
            )

            if (
                avg_value is not None
                and avg_value < 0.50
            ):

                cell.fill = blue_fill

                output_ws.cell(
                    row=row,
                    column=1
                ).fill = blue_fill

                condition_count[row] += 1

    # ========================================================
    # CONDITION 3
    #
    # 16< D-B / Avg.4
    #
    # Parentheses < first value
    # AND
    # Parentheses < -1.00
    # ========================================================

    if col_16lt:

        for row in range(
            2,
            output_ws.max_row + 1
        ):

            cell = output_ws.cell(
                row=row,
                column=col_16lt
            )

            first_val = first_number(
                cell.value
            )

            avg_val = (
                number_inside_parentheses(
                    cell.value
                )
            )

            if (
                first_val is not None
                and avg_val is not None
                and avg_val < first_val
                and avg_val < -1.00
            ):

                cell.fill = blue_fill

                output_ws.cell(
                    row=row,
                    column=1
                ).fill = blue_fill

                condition_count[row] += 1

    # ========================================================
    # CONDITION 4
    #
    # Sum O2H.10 BELOW AVERAGE
    # ========================================================

    o2h_values = []

    if col_o2h:

        for row in range(
            2,
            output_ws.max_row + 1
        ):

            value = numeric_value(
                output_ws.cell(
                    row=row,
                    column=col_o2h
                ).value
            )

            if value is not None:

                o2h_values.append(
                    value
                )

    o2h_average = (
        sum(o2h_values)
        / len(o2h_values)
        if o2h_values
        else None
    )

    if (
        col_o2h
        and o2h_average is not None
    ):

        for row in range(
            2,
            output_ws.max_row + 1
        ):

            value = numeric_value(
                output_ws.cell(
                    row=row,
                    column=col_o2h
                ).value
            )

            if (
                value is not None
                and value < o2h_average
            ):

                output_ws.cell(
                    row=row,
                    column=col_o2h
                ).fill = blue_fill

                output_ws.cell(
                    row=row,
                    column=1
                ).fill = blue_fill

                condition_count[row] += 1

    # ========================================================
    # CONDITION 5
    #
    # Sum O2L.10 BELOW AVERAGE
    # ========================================================

    o2l_values = []

    if col_o2l:

        for row in range(
            2,
            output_ws.max_row + 1
        ):

            value = numeric_value(
                output_ws.cell(
                    row=row,
                    column=col_o2l
                ).value
            )

            if value is not None:

                o2l_values.append(
                    value
                )

    o2l_average = (
        sum(o2l_values)
        / len(o2l_values)
        if o2l_values
        else None
    )

    if (
        col_o2l
        and o2l_average is not None
    ):

        for row in range(
            2,
            output_ws.max_row + 1
        ):

            value = numeric_value(
                output_ws.cell(
                    row=row,
                    column=col_o2l
                ).value
            )

            if (
                value is not None
                and value < o2l_average
            ):

                output_ws.cell(
                    row=row,
                    column=col_o2l
                ).fill = blue_fill

                output_ws.cell(
                    row=row,
                    column=1
                ).fill = blue_fill

                condition_count[row] += 1

    # ========================================================
    # CONDITION 6
    #
    # MOST RECENT P2L/O2L < -1.00
    # ========================================================

    recent_heading = None

    if col_recent_p2l:

        recent_heading = (
            output_ws.cell(
                row=1,
                column=col_recent_p2l
            ).value
        )

        for row in range(
            2,
            output_ws.max_row + 1
        ):

            value = numeric_value(
                output_ws.cell(
                    row=row,
                    column=col_recent_p2l
                ).value
            )

            if (
                value is not None
                and value < -1.00
            ):

                output_ws.cell(
                    row=row,
                    column=col_recent_p2l
                ).fill = blue_fill

                output_ws.cell(
                    row=row,
                    column=1
                ).fill = blue_fill

                condition_count[row] += 1

    # ========================================================
    # STRONGEST ROW(S)
    # ========================================================

    valid_counts = [
        count
        for count in condition_count.values()
        if count > 0
    ]

    if valid_counts:

        maximum_count = max(
            valid_counts
        )

        strongest_rows = [
            row
            for row, count
            in condition_count.items()
            if count == maximum_count
        ]

        for row in strongest_rows:

            set_row_green(
                output_ws,
                row,
                green_fill
            )

    else:

        maximum_count = 0
        strongest_rows = []

    # ========================================================
    # COLUMN WIDTH
    # ========================================================

    for col in range(
        1,
        output_ws.max_column + 1
    ):

        max_length = 0

        for row in range(
            1,
            min(
                output_ws.max_row,
                1000
            ) + 1
        ):

            value = output_ws.cell(
                row=row,
                column=col
            ).value

            if value is not None:

                length = len(
                    str(value)
                )

                if length > max_length:
                    max_length = length

        output_ws.column_dimensions[
            get_column_letter(col)
        ].width = min(
            max(
                max_length + 2,
                10
            ),
            35
        )

    # ========================================================
    # SAVE TO MEMORY
    # ========================================================

    output_buffer = io.BytesIO()

    output_wb.save(
        output_buffer
    )

    output_wb.close()
    source_wb.close()

    output_buffer.seek(0)

    return (
        output_buffer.getvalue(),
        {
            "o2h_average": o2h_average,
            "o2l_average": o2l_average,
            "recent_heading": recent_heading,
            "maximum_count": maximum_count,
            "strongest_rows": strongest_rows
        }
    )


# ============================================================
# GENERATE OUTPUT
# ============================================================

def generate_output():

    master_bytes = download_master_bytes()

    if not master_bytes:

        raise ValueError(
            "No Master file is currently available."
        )

    output_bytes, result = create_output(
        master_bytes
    )

    return output_bytes, result


# ============================================================
# TELEGRAM
# ============================================================

def send_telegram_message(
    message
):

    settings = get_settings()

    if not settings.get(
        "telegram_enabled",
        False
    ):

        return False, (
            "Telegram is disabled."
        )

    bot_token = safe_text(
        settings.get(
            "telegram_bot_token",
            ""
        )
    )

    chat_id = safe_text(
        settings.get(
            "telegram_chat_id",
            ""
        )
    )

    if not bot_token or not chat_id:

        return False, (
            "Telegram Bot Token or Chat ID is missing."
        )

    url = (
        "https://api.telegram.org/bot"
        f"{bot_token}/sendMessage"
    )

    payload = {
        "chat_id": chat_id,
        "text": message
    }

    try:

        response = requests.post(
            url,
            data=payload,
            timeout=15
        )

        if response.ok:

            return True, "Sent"

        return False, response.text

    except Exception as e:

        return False, str(e)


# ============================================================
# ADMIN USER FUNCTIONS
# ============================================================

def list_users():

    require_admin()

    try:

        response = (
            get_admin_client()
            .table("profiles")
            .select(
                "id,email,role,is_active,created_at"
            )
            .order(
                "created_at",
                desc=False
            )
            .execute()
        )

        return response.data or []

    except Exception as e:

        st.error(
            f"Unable to load users: {e}"
        )

        return []


def create_user(
    email,
    password,
    role="user"
):

    require_admin()

    email = email.strip().lower()

    if not email:
        return False, "Email is required."

    if len(password) < 8:

        return False, (
            "Password must contain at least 8 characters."
        )

    if role not in [
        "admin",
        "user"
    ]:

        role = "user"

    try:

        response = (
            get_admin_client()
            .auth.admin.create_user(
                {
                    "email": email,
                    "password": password,
                    "email_confirm": True,
                    "user_metadata": {
                        "app_role": role
                    }
                }
            )
        )

        new_user = response.user

        if not new_user:

            return False, (
                "User creation failed."
            )

        # Explicitly set role in profile
        get_admin_client().table(
            "profiles"
        ).upsert(
            {
                "id": new_user.id,
                "email": email,
                "role": role,
                "is_active": True
            }
        ).execute()

        return True, (
            f"User {email} created successfully."
        )

    except Exception as e:

        return False, str(e)


def update_user_password(
    user_id,
    new_password
):

    require_admin()

    if len(new_password) < 8:

        return False, (
            "Password must contain at least 8 characters."
        )

    try:

        get_admin_client().auth.admin.update_user_by_id(
            user_id,
            {
                "password": new_password
            }
        )

        return True, (
            "Password changed successfully."
        )

    except Exception as e:

        return False, str(e)


def update_user_role(
    user_id,
    role
):

    require_admin()

    if role not in [
        "admin",
        "user"
    ]:

        return False, (
            "Invalid role."
        )

    try:

        # Prevent accidental self-demotion
        current_user = st.session_state.user

        if (
            current_user
            and current_user.id == user_id
            and role != "admin"
        ):

            return False, (
                "You cannot remove your own admin access."
            )

        get_admin_client().table(
            "profiles"
        ).update(
            {
                "role": role
            }
        ).eq(
            "id",
            user_id
        ).execute()

        return True, (
            "Role updated."
        )

    except Exception as e:

        return False, str(e)


def update_user_active(
    user_id,
    active
):

    require_admin()

    current_user = st.session_state.user

    if (
        current_user
        and current_user.id == user_id
        and not active
    ):

        return False, (
            "You cannot disable your own account."
        )

    try:

        get_admin_client().table(
            "profiles"
        ).update(
            {
                "is_active": bool(active)
            }
        ).eq(
            "id",
            user_id
        ).execute()

        return True, (
            "User status updated."
        )

    except Exception as e:

        return False, str(e)


def delete_user(
    user_id
):

    require_admin()

    current_user = st.session_state.user

    if (
        current_user
        and current_user.id == user_id
    ):

        return False, (
            "You cannot delete your own account."
        )

    try:

        get_admin_client().auth.admin.delete_user(
            user_id
        )

        return True, (
            "User deleted successfully."
        )

    except Exception as e:

        return False, str(e)


# ============================================================
# SAVE SETTINGS
# ============================================================

def save_settings(
    values
):

    require_admin()

    try:

        get_admin_client().table(
            "app_settings"
        ).upsert(
            {
                "id": 1,
                **values
            }
        ).execute()

        return True, (
            "Settings saved successfully."
        )

    except Exception as e:

        return False, str(e)


# ============================================================
# LOGIN PAGE
# ============================================================

def show_login():

    st.markdown(
        """
        <style>
        .login-title {
            text-align:center;
            font-size:32px;
            font-weight:700;
            margin-top:40px;
        }

        .login-subtitle {
            text-align:center;
            color:#777;
            margin-bottom:25px;
        }
        </style>
        """,
        unsafe_allow_html=True
    )

    st.markdown(
        f"""
        <div class="login-title">
            {APP_NAME}
        </div>

        <div class="login-subtitle">
            Secure Login
        </div>
        """,
        unsafe_allow_html=True
    )

    left, center, right = st.columns(
        [1, 2, 1]
    )

    with center:

        with st.form(
            "login_form"
        ):

            email = st.text_input(
                "Email",
                autocomplete="email"
            )

            password = st.text_input(
                "Password",
                type="password",
                autocomplete="current-password"
            )

            login_button = st.form_submit_button(
                "Login",
                use_container_width=True,
                type="primary"
            )

            if login_button:

                if not email or not password:

                    st.error(
                        "Please enter email and password."
                    )

                else:

                    with st.spinner(
                        "Signing in..."
                    ):

                        success, message = login_user(
                            email,
                            password
                        )

                    if success:

                        st.success(
                            message
                        )

                        st.rerun()

                    else:

                        st.error(
                            message
                        )

        st.caption(
            "User registration is disabled. "
            "Accounts are created by the administrator."
        )


# ============================================================
# SIDEBAR
# ============================================================

def show_sidebar():

    user = st.session_state.user

    with st.sidebar:

        st.markdown(
            f"### {APP_NAME}"
        )

        st.divider()

        st.write(
            f"**User:** {user.email}"
        )

        if is_admin():

            st.success(
                "ADMIN ACCESS"
            )

        else:

            st.info(
                "USER ACCESS — Read-only controls"
            )

        st.divider()

        metadata = get_master_metadata()

        if metadata:

            st.success(
                "Master: Available"
            )

            filename = metadata.get(
                "original_filename",
                "master.xlsx"
            )

            st.caption(
                f"File: {filename}"
            )

            uploaded_at = metadata.get(
                "uploaded_at"
            )

            if uploaded_at:

                st.caption(
                    f"Uploaded: {uploaded_at}"
                )

        else:

            st.warning(
                "Master: Not uploaded"
            )

        st.divider()

        if st.button(
            "Logout",
            use_container_width=True
        ):

            logout_user()


# ============================================================
# MASTER MANAGEMENT - ADMIN
# ============================================================

def admin_master_panel():

    require_admin()

    st.subheader(
        "Master File Management"
    )

    metadata = get_master_metadata()

    if metadata:

        col1, col2, col3 = st.columns(3)

        with col1:

            st.metric(
                "Master Status",
                "Available"
            )

        with col2:

            st.metric(
                "File Size",
                f"{metadata.get('file_size', 0) / 1024:.1f} KB"
            )

        with col3:

            st.metric(
                "Original File",
                metadata.get(
                    "original_filename",
                    "-"
                )
            )

    else:

        st.info(
            "No Master file is currently stored."
        )

    uploaded_file = st.file_uploader(
        "Upload / Replace Master Excel",
        type=[
            "xlsx",
            "xlsm",
            "xltx",
            "xltm"
        ],
        key="master_upload"
    )

    col1, col2 = st.columns(2)

    with col1:

        if st.button(
            "Upload / Replace Master",
            type="primary",
            use_container_width=True,
            disabled=uploaded_file is None
        ):

            with st.spinner(
                "Uploading Master..."
            ):

                success, message = save_master_file(
                    uploaded_file,
                    st.session_state.user.id
                )

            if success:

                st.success(
                    message
                )

                st.rerun()

            else:

                st.error(
                    message
                )

    with col2:

        delete_confirm = st.checkbox(
            "Confirm Master deletion",
            key="confirm_master_delete"
        )

        if st.button(
            "Delete Master",
            type="secondary",
            use_container_width=True,
            disabled=not delete_confirm
        ):

            success, message = delete_master()

            if success:

                st.success(
                    message
                )

                st.rerun()

            else:

                st.error(
                    message
                )


# ============================================================
# VIEW MASTER SUMMARY
# ============================================================

def show_master_view():

    st.subheader(
        "Master Summary"
    )

    master_bytes = download_master_bytes()

    if not master_bytes:

        st.warning(
            "No Master file is currently available."
        )

        return

    try:

        wb = load_workbook(
            io.BytesIO(master_bytes),
            data_only=True,
            read_only=True
        )

        ws = wb["summary"]

        rows = ws.iter_rows(
            values_only=True
        )

        rows = list(rows)

        wb.close()

        if not rows:

            st.warning(
                "Summary sheet is empty."
            )

            return

        header = rows[0]

        data = rows[1:]

        # Make unique column names
        columns = []

        used = {}

        for col in header:

            name = (
                str(col)
                if col is not None
                else "Column"
            )

            if name in used:

                used[name] += 1

                name = (
                    f"{name}_{used[name]}"
                )

            else:

                used[name] = 0

            columns.append(name)

        df = pd.DataFrame(
            data,
            columns=columns
        )

        st.caption(
            f"{len(df):,} rows × {len(df.columns):,} columns"
        )

        st.dataframe(
            df,
            use_container_width=True,
            height=600,
            hide_index=True
        )

    except Exception as e:

        st.error(
            f"Unable to display Master summary: {e}"
        )


# ============================================================
# GENERATE / DOWNLOAD
# ============================================================

def generate_download_panel():

    st.subheader(
        "Generate Value-Only Output"
    )

    metadata = get_master_metadata()

    if not metadata:

        st.warning(
            "Admin has not uploaded a Master file yet."
        )

        return

    st.info(
        "The output is generated from the current "
        "Master file using the original 6thSense conditions."
    )

    if st.button(
        "Generate & Download",
        type="primary",
        use_container_width=True
    ):

        with st.spinner(
            "Generating Excel output..."
        ):

            try:

                output_bytes, result = (
                    generate_output()
                )

                st.session_state.generated_output = (
                    output_bytes
                )

                st.session_state.generated_result = (
                    result
                )

                st.success(
                    "Output generated successfully."
                )

            except Exception as e:

                st.error(
                    f"Generation failed: {e}"
                )

    if (
        "generated_output"
        in st.session_state
    ):

        output_bytes = (
            st.session_state.generated_output
        )

        result = (
            st.session_state.get(
                "generated_result",
                {}
            )
        )

        st.download_button(
            "Download Excel",
            data=output_bytes,
            file_name=OUTPUT_NAME,
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            use_container_width=True
        )

        col1, col2, col3 = st.columns(3)

        with col1:

            if result.get(
                "o2h_average"
            ) is not None:

                st.metric(
                    "Sum O2H.10 Average",
                    f"{result['o2h_average']:.4f}"
                )

        with col2:

            if result.get(
                "o2l_average"
            ) is not None:

                st.metric(
                    "Sum O2L.10 Average",
                    f"{result['o2l_average']:.4f}"
                )

        with col3:

            st.metric(
                "Highest Condition Count",
                result.get(
                    "maximum_count",
                    0
                )
            )

        if result.get(
            "recent_heading"
        ):

            st.caption(
                "Recent P2L/O2L heading: "
                f"{result['recent_heading']}"
            )


# ============================================================
# ADMIN USER MANAGEMENT
# ============================================================

def admin_user_panel():

    require_admin()

    st.subheader(
        "User Management"
    )

    users = list_users()

    if users:

        display_users = []

        for user in users:

            display_users.append(
                {
                    "Email": user.get(
                        "email",
                        ""
                    ),
                    "Role": user.get(
                        "role",
                        "user"
                    ),
                    "Active": user.get(
                        "is_active",
                        True
                    ),
                    "Created": user.get(
                        "created_at",
                        ""
                    )
                }
            )

        st.dataframe(
            pd.DataFrame(
                display_users
            ),
            use_container_width=True,
            hide_index=True
        )

    st.divider()

    st.markdown(
        "### Create User"
    )

    with st.form(
        "create_user_form"
    ):

        new_email = st.text_input(
            "Email"
        )

        new_password = st.text_input(
            "Initial Password",
            type="password"
        )

        new_role = st.selectbox(
            "Role",
            [
                "user",
                "admin"
            ]
        )

        create_button = st.form_submit_button(
            "Create User",
            type="primary"
        )

        if create_button:

            success, message = create_user(
                new_email,
                new_password,
                new_role
            )

            if success:

                st.success(
                    message
                )

                st.rerun()

            else:

                st.error(
                    message
                )

    st.divider()

    if users:

        st.markdown(
            "### Manage Existing User"
        )

        user_options = {
            f"{u.get('email')} ({u.get('role')})":
                u
            for u in users
        }

        selected_label = st.selectbox(
            "Select User",
            list(
                user_options.keys()
            )
        )

        selected_user = user_options[
            selected_label
        ]

        selected_id = selected_user[
            "id"
        ]

        col1, col2 = st.columns(2)

        with col1:

            new_role = st.selectbox(
                "Change Role",
                [
                    "user",
                    "admin"
                ],
                index=(
                    0
                    if selected_user.get(
                        "role"
                    ) == "user"
                    else 1
                )
            )

            if st.button(
                "Save Role",
                use_container_width=True
            ):

                success, message = (
                    update_user_role(
                        selected_id,
                        new_role
                    )
                )

                if success:

                    st.success(
                        message
                    )

                    st.rerun()

                else:

                    st.error(
                        message
                    )

        with col2:

            active = st.checkbox(
                "Account Active",
                value=bool(
                    selected_user.get(
                        "is_active",
                        True
                    )
                )
            )

            if st.button(
                "Save Account Status",
                use_container_width=True
            ):

                success, message = (
                    update_user_active(
                        selected_id,
                        active
                    )
                )

                if success:

                    st.success(
                        message
                    )

                    st.rerun()

                else:

                    st.error(
                        message
                    )

        new_password = st.text_input(
            "New Password",
            type="password",
            key="admin_change_password"
        )

        if st.button(
            "Change Password",
            use_container_width=True
        ):

            if not new_password:

                st.error(
                    "Enter a new password."
                )

            else:

                success, message = (
                    update_user_password(
                        selected_id,
                        new_password
                    )
                )

                if success:

                    st.success(
                        message
                    )

                else:

                    st.error(
                        message
                    )

        st.divider()

        confirm_delete = st.checkbox(
            "Confirm permanent user deletion",
            key="confirm_user_delete"
        )

        if st.button(
            "Delete Selected User",
            type="secondary",
            use_container_width=True,
            disabled=not confirm_delete
        ):

            success, message = delete_user(
                selected_id
            )

            if success:

                st.success(
                    message
                )

                st.rerun()

            else:

                st.error(
                    message
                )


# ============================================================
# ADMIN SETTINGS
# ============================================================

def admin_settings_panel():

    require_admin()

    st.subheader(
        "Application Settings"
    )

    settings = get_settings()

    st.markdown(
        "### Excel Colors"
    )

    col1, col2 = st.columns(2)

with col1:

    light_blue = st.text_input(
        "Light Blue HEX",
            value=settings.get(
                "light_blue",
                "ADD8E6"
            )
        )

    with col2:

        light_green = st.text_input(
            "Light Green HEX",
            value=settings.get(
                "light_green",
                "90EE90"
            )
        )

    st.markdown(
        "### Telegram"
    )

    telegram_enabled = st.checkbox(
        "Enable Telegram",
        value=settings.get(
            "telegram_enabled",
            False
        )
    )

    telegram_token = st.text_input(
        "Telegram Bot Token",
        value=settings.get(
            "telegram_bot_token",
            ""
        ),
        type="password"
    )

    telegram_chat_id = st.text_input(
        "Telegram Chat ID",
        value=settings.get(
            "telegram_chat_id",
            ""
        )
    )

    st.markdown(
        "### Alert Condition"
    )

    alert_enabled = st.checkbox(
        "Enable O2L Alert",
        value=settings.get(
            "alert_enabled",
            False
        )
    )

    alert_operator = st.selectbox(
        "Operator",
        [
            "<",
            "<=",
            ">",
            ">=",
            "="
        ],
        index=[
            "<",
            "<=",
            ">",
            ">=",
            "="
        ].index(
            settings.get(
                "alert_operator",
                "<"
            )
        )
    )

    alert_value = st.number_input(
        "O2L Alert Value",
        value=float(
            settings.get(
                "alert_value",
                -1.00
            )
        ),
        step=0.10
    )

    if st.button(
        "Save Settings",
        type="primary"
    ):

        # Basic HEX validation
        blue = light_blue.replace(
            "#",
            ""
        ).upper()

        green = light_green.replace(
            "#",
            ""
        ).upper()

        if not re.fullmatch(
            r"[0-9A-F]{6}",
            blue
        ):

            st.error(
                "Invalid Light Blue HEX value."
            )

        elif not re.fullmatch(
            r"[0-9A-F]{6}",
            green
        ):

            st.error(
                "Invalid Light Green HEX value."
            )

        else:

            success, message = save_settings(
                {
                    "light_blue": blue,
                    "light_green": green,
                    "telegram_enabled": telegram_enabled,
                    "telegram_bot_token": telegram_token,
                    "telegram_chat_id": telegram_chat_id,
                    "alert_enabled": alert_enabled,
                    "alert_type": "O2L",
                    "alert_operator": alert_operator,
                    "alert_value": float(alert_value)
                }
            )

            if success:

                st.success(
                    message
                )

            else:

                st.error(
                    message
                )

     st.divider()

    if st.button(
        "Send Telegram Test"
    ):

        success, message = send_telegram_message(
            "6thSense Trading\n"
            "Telegram test alert successful."
        )

        if success:

            st.success(
                message
            )

        else:

            st.error(
                message
            )


# ============================================================
# ADMIN DASHBOARD
# ============================================================

def show_admin_dashboard():

    require_admin()

    st.title(
        APP_NAME
    )

    st.success(
        "ADMIN ACCESS"
    )

    tabs = st.tabs(
        [
            "📁 Master",
            "📊 View",
            "⬇️ Generate",
            "👥 Users",
            "⚙️ Settings"
        ]
    )

    with tabs[0]:

        admin_master_panel()

    with tabs[1]:

        show_master_view()

    with tabs[2]:

        generate_download_panel()

    with tabs[3]:

        admin_user_panel()

    with tabs[4]:

        admin_settings_panel()


# ============================================================
# USER DASHBOARD
# ============================================================

def show_user_dashboard():

    st.title(
        APP_NAME
    )

    st.info(
        "USER ACCESS — View, Generate and Download only."
    )

    tabs = st.tabs(
        [
            "📊 View Master",
            "⬇️ Generate & Download"
        ]
    )

    with tabs[0]:

        show_master_view()

    with tabs[1]:

        generate_download_panel()


# ============================================================
# APPLICATION START
# ============================================================

if not refresh_session():

    # Clear stale session
    st.session_state.access_token = None
    st.session_state.refresh_token = None
    st.session_state.user = None
    st.session_state.role = None
    st.session_state.login_email = None

    show_login()

    st.stop()


# ============================================================
# LOGGED-IN USER
# ============================================================

show_sidebar()


# ============================================================
# ROLE-BASED APPLICATION
# ============================================================

if is_admin():

    show_admin_dashboard()

else:

    show_user_dashboard()




  



      

                      

                      


      




          

          

      



              

              





          





