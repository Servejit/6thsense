# ============================================================
# 6thSense (6S-FO200) Vardaan
# STREAMLIT + SUPABASE SECURE APPLICATION
#
# PART 1 OF 5
#
# This part contains:
#   - Imports
#   - Streamlit configuration
#   - Supabase configuration
#   - Security constants
#   - Session-state initialization
#   - Supabase helper functions
#   - Authentication helpers
#   - Role helpers
#   - Safe utility functions
#
# PART 2 WILL CONTINUE DIRECTLY AFTER THIS CODE
# ============================================================


# ============================================================
#                    IMPORTS
# ============================================================

import streamlit as st

import os
import io
import re
import time
import json
import hashlib
import traceback
from datetime import datetime, timezone

import pandas as pd
import numpy as np

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from supabase import create_client, Client


# ============================================================
#                    STREAMLIT CONFIG
# ============================================================

st.set_page_config(
    page_title="6thSense (6S-FO200) Vardaan",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ============================================================
#                    APPLICATION CONSTANTS
# ============================================================

APP_NAME = "6thSense (6S-FO200) Vardaan"

APP_VERSION = "1.0.0"

MASTER_FILE_NAME = "master.xlsx"

OUTPUT_FILE_NAME = "summary_output.xlsx"

MAX_UPLOAD_MB = 100

ALLOWED_EXTENSIONS = {
    ".xlsx",
    ".xlsm",
}

ADMIN_ROLE = "admin"

USER_ROLE = "user"


# ============================================================
#                    SUPABASE CONFIG
# ============================================================
#
# IMPORTANT:
#
# Put these two values in Streamlit Cloud Secrets.
#
# Example:
#
# [supabase]
# url = "https://YOUR_PROJECT.supabase.co"
# key = "YOUR_SUPABASE_ANON_KEY"
#
# Do NOT put your service_role key into this application.
#
# The anon key is intended for client-side/public applications
# when Row Level Security (RLS) is correctly configured.
#
# ============================================================


def get_supabase_config():
    """
    Read Supabase configuration from Streamlit secrets first,
    then environment variables as a fallback.
    """

    url = ""
    key = ""

    # --------------------------------------------------------
    # Streamlit secrets
    # --------------------------------------------------------

    try:

        if "supabase" in st.secrets:

            supabase_section = st.secrets["supabase"]

            url = str(
                supabase_section.get(
                    "url",
                    ""
                )
            ).strip()

            key = str(
                supabase_section.get(
                    "key",
                    ""
                )
            ).strip()

    except Exception:
        pass


    # --------------------------------------------------------
    # Environment-variable fallback
    # --------------------------------------------------------

    if not url:

        url = os.environ.get(
            "SUPABASE_URL",
            ""
        ).strip()


    if not key:

        key = os.environ.get(
            "SUPABASE_KEY",
            ""
        ).strip()


    return url, key


# ============================================================
#                    CREATE SUPABASE CLIENT
# ============================================================


SUPABASE_URL, SUPABASE_KEY = get_supabase_config()


supabase: Client | None = None


if SUPABASE_URL and SUPABASE_KEY:

    try:

        supabase = create_client(
            SUPABASE_URL,
            SUPABASE_KEY
        )

    except Exception as e:

        supabase = None

        st.error(
            "Unable to initialize the Supabase connection."
        )

        st.stop()


# ============================================================
#                    BASIC SECURITY HELPERS
# ============================================================


def clean_text(
    value,
    max_length=500
):
    """
    Safely convert a value to a trimmed string.
    """

    if value is None:

        return ""

    try:

        value = str(value)

    except Exception:

        return ""

    value = value.strip()

    if len(value) > max_length:

        value = value[:max_length]

    return value


# ============================================================
#                    EMAIL VALIDATION
# ============================================================


def is_valid_email(
    email
):

    email = clean_text(
        email,
        320
    ).lower()

    if not email:

        return False

    pattern = (
        r"^[A-Za-z0-9.!#$%&'*+/=?^_`{|}~-]+"
        r"@[A-Za-z0-9-]+"
        r"(?:\.[A-Za-z0-9-]+)+$"
    )

    return bool(
        re.match(
            pattern,
            email
        )
    )


# ============================================================
#                    PASSWORD SAFETY CHECK
# ============================================================


def password_is_reasonable(
    password
):

    if password is None:

        return False

    password = str(
        password
    )

    if len(password) < 6:

        return False

    return True


# ============================================================
#                    FILE VALIDATION
# ============================================================


def validate_excel_upload(
    uploaded_file
):

    if uploaded_file is None:

        return False, "No file selected."

    filename = clean_text(
        uploaded_file.name,
        255
    )

    extension = os.path.splitext(
        filename
    )[1].lower()

    if extension not in ALLOWED_EXTENSIONS:

        return (
            False,
            "Only .xlsx or .xlsm Excel files are allowed."
        )


    try:

        file_size = uploaded_file.size

    except Exception:

        file_size = 0


    maximum_bytes = (
        MAX_UPLOAD_MB
        * 1024
        * 1024
    )


    if file_size > maximum_bytes:

        return (
            False,
            f"File is larger than {MAX_UPLOAD_MB} MB."
        )


    return True, ""


# ============================================================
#                    FILE HASH
# ============================================================


def calculate_file_hash(
    file_bytes
):

    if not file_bytes:

        return ""


    return hashlib.sha256(
        file_bytes
    ).hexdigest()


# ============================================================
#                    CURRENT UTC TIME
# ============================================================


def utc_now():

    return datetime.now(
        timezone.utc
    ).isoformat()


# ============================================================
#                    SUPABASE AVAILABILITY
# ============================================================


def require_supabase():

    if supabase is None:

        st.error(
            """
            Supabase is not configured.

            Please configure SUPABASE_URL and SUPABASE_KEY
            in Streamlit Secrets.
            """
        )

        st.stop()


# ============================================================
#                    SESSION STATE INITIALIZATION
# ============================================================


def initialize_session_state():

    defaults = {

        "user": None,

        "session": None,

        "is_admin": False,

        "profile": None,

        "authenticated": False,

        "master_file_exists": False,

        "master_file_name": MASTER_FILE_NAME,

        "generated_output": None,

        "generated_output_name": OUTPUT_FILE_NAME,

        "last_error": None,

        "login_attempts": 0,

        "last_login_time": 0.0,

        "app_initialized": True,

    }


    for key, value in defaults.items():

        if key not in st.session_state:

            st.session_state[key] = value


# Initialize immediately.

initialize_session_state()


# ============================================================
#                    SUPABASE AUTH HELPERS
# ============================================================


def get_current_user():

    """
    Return the currently authenticated Supabase user.

    Returns:
        User object or None.
    """

    require_supabase()

    try:

        response = supabase.auth.get_user()

        if response is None:

            return None

        user = getattr(
            response,
            "user",
            None
        )

        return user

    except Exception:

        return None


# ============================================================
#                    SUPABASE SESSION
# ============================================================


def get_current_session():

    """
    Return current Supabase authentication session.
    """

    require_supabase()

    try:

        response = supabase.auth.get_session()

        return response

    except Exception:

        return None


# ============================================================
#                    SIGN IN
# ============================================================


def sign_in_user(
    email,
    password
):

    """
    Authenticate a user using Supabase Auth.
    """

    require_supabase()

    email = clean_text(
        email,
        320
    ).lower()

    password = str(
        password
    )


    if not is_valid_email(
        email
    ):

        return (
            False,
            None,
            None,
            "Please enter a valid email address."
        )


    if not password_is_reasonable(
        password
    ):

        return (
            False,
            None,
            None,
            "Invalid password."
        )


    try:

        response = supabase.auth.sign_in_with_password(
            {
                "email": email,
                "password": password,
            }
        )


        if response is None:

            return (
                False,
                None,
                None,
                "Login failed."
            )


        session = getattr(
            response,
            "session",
            None
        )

        user = getattr(
            response,
            "user",
            None
        )


        if user is None:

            return (
                False,
                None,
                None,
                "Invalid email or password."
            )


        return (
            True,
            user,
            session,
            None
        )


    except Exception as e:

        error_text = clean_text(
            str(e),
            1000
        )


        # Do not expose internal database details
        # to the user.

        lowered = error_text.lower()


        if (
            "invalid login credentials"
            in lowered
        ):

            message = (
                "Invalid email or password."
            )

        elif (
            "email not confirmed"
            in lowered
        ):

            message = (
                "Please confirm your email before logging in."
            )

        else:

            message = (
                "Login failed. Please try again."
            )


        return (
            False,
            None,
            None,
            message
        )


# ============================================================
#                    SIGN OUT
# ============================================================


def sign_out_user():

    """
    Safely sign out the current Supabase user and clear
    local Streamlit authentication state.
    """

    try:

        if supabase is not None:

            supabase.auth.sign_out()

    except Exception:

        # Continue clearing local state even if the
        # remote logout encounters an error.

        pass


    # --------------------------------------------------------
    # Clear application authentication state
    # --------------------------------------------------------

    st.session_state.user = None

    st.session_state.session = None

    st.session_state.profile = None

    st.session_state.is_admin = False

    st.session_state.authenticated = False

    st.session_state.generated_output = None

    st.session_state.last_error = None


# ============================================================
#                    USER ID HELPER
# ============================================================


def get_user_id(
    user=None
):

    if user is None:

        user = st.session_state.get(
            "user"
        )


    if user is None:

        return None


    try:

        user_id = getattr(
            user,
            "id",
            None
        )

    except Exception:

        return None


    if not user_id:

        return None


    return str(
        user_id
    )


# ============================================================
#                    USER EMAIL HELPER
# ============================================================


def get_user_email(
    user=None
):

    if user is None:

        user = st.session_state.get(
            "user"
        )


    if user is None:

        return ""


    try:

        email = getattr(
            user,
            "email",
            ""
        )

    except Exception:

        return ""


    return clean_text(
        email,
        320
    ).lower()


# ============================================================
#                    PROFILE LOOKUP
# ============================================================
#
# Expected Supabase table:
#
# profiles
#
# Suggested columns:
#
# id          uuid
# email       text
# role        text
# is_active   boolean
#
# The SQL/RLS setup will be provided with the complete project
# configuration after the five app.py parts.
#
# ============================================================


def get_user_profile(
    user_id
):

    require_supabase()

    if not user_id:

        return None


    try:

        response = (
            supabase
            .table("profiles")
            .select(
                "id,email,role,is_active"
            )
            .eq(
                "id",
                user_id
            )
            .limit(1)
            .execute()
        )


        data = getattr(
            response,
            "data",
            None
        )


        if not data:

            return None


        return data[0]


    except Exception:

        return None


# ============================================================
#                    ROLE CHECK
# ============================================================


def user_is_admin(
    user_id=None
):

    """
    Server-side role lookup through Supabase.

    Do not rely only on a Streamlit checkbox or URL parameter
    to decide whether a person is an administrator.
    """

    if user_id is None:

        user_id = get_user_id()


    if not user_id:

        return False


    profile = get_user_profile(
        user_id
    )


    if not profile:

        return False


    role = clean_text(
        profile.get(
            "role",
            ""
        ),
        50
    ).lower()


    is_active = profile.get(
        "is_active",
        True
    )


    if not is_active:

        return False


    return (
        role == ADMIN_ROLE
    )


# ============================================================
#                    LOAD AUTHENTICATED USER
# ============================================================


def refresh_authenticated_user():

    """
    Refresh authentication information from Supabase.

    This is intentionally done through Supabase rather than
    trusting values stored only in Streamlit session state.
    """

    require_supabase()

    user = get_current_user()


    if user is None:

        st.session_state.user = None

        st.session_state.session = None

        st.session_state.profile = None

        st.session_state.is_admin = False

        st.session_state.authenticated = False

        return False


    user_id = get_user_id(
        user
    )


    if not user_id:

        sign_out_user()

        return False


    profile = get_user_profile(
        user_id
    )


    # --------------------------------------------------------
    # A valid Auth account without an active application
    # profile is not allowed into the application.
    # --------------------------------------------------------

    if profile is None:

        sign_out_user()

        st.session_state.last_error = (
            "Your account is not authorized for this application."
        )

        return False


    is_active = profile.get(
        "is_active",
        True
    )


    if not is_active:

        sign_out_user()

        st.session_state.last_error = (
            "Your account has been disabled."
        )

        return False


    role = clean_text(
        profile.get(
            "role",
            USER_ROLE
        ),
        50
    ).lower()


    if role not in {
        ADMIN_ROLE,
        USER_ROLE,
    }:

        role = USER_ROLE


    # --------------------------------------------------------
    # Store only validated information.
    # --------------------------------------------------------

    st.session_state.user = user

    st.session_state.session = (
        get_current_session()
    )

    st.session_state.profile = profile

    st.session_state.is_admin = (
        role == ADMIN_ROLE
    )

    st.session_state.authenticated = True


    return True


# ============================================================
#                    SAFE ERROR HANDLER
# ============================================================


def show_safe_error(
    message,
    details=None
):

    """
    Show a user-friendly error without exposing secrets,
    access tokens, passwords, or raw database internals.
    """

    message = clean_text(
        message,
        1000
    )


    if not message:

        message = (
            "An unexpected error occurred."
        )


    st.error(
        message
    )


    # Detailed traceback is deliberately NOT displayed to the
    # normal user.
    #
    # If debugging is needed during development, use the
    # server logs rather than exposing internal details.


# ============================================================
#                    APPLICATION HEADER
# ============================================================


def show_application_header():

    st.title(
        APP_NAME
    )

    st.caption(
        "Secure Streamlit + Supabase application"
    )


# ============================================================
#                    INITIAL SUPABASE CHECK
# ============================================================


if supabase is None:

    st.title(
        APP_NAME
    )

    st.error(
        """
        Supabase configuration is missing.

        Configure the following Streamlit Secrets:

        [supabase]
        url = "YOUR_SUPABASE_PROJECT_URL"
        key = "YOUR_SUPABASE_ANON_KEY"
        """
    )

    st.stop()


# ============================================================
#                    END OF PART 1
# ============================================================
# ============================================================
# 6thSense (6S-FO200) Vardaan
# STREAMLIT + SUPABASE SECURE APPLICATION
#
# PART 2 OF 5
#
# Contains:
#   - Authentication startup
#   - Login screen
#   - Login protection
#   - Logout
#   - Sidebar
#   - Admin/User role display
#   - Session handling
#
# CONTINUES DIRECTLY FROM PART 1
# ============================================================


# ============================================================
#                    AUTHENTICATION STARTUP
# ============================================================

def initialize_authentication():

    """
    Check whether a valid Supabase authentication session
    already exists.

    Returns:
        True  -> authenticated
        False -> not authenticated
    """

    # --------------------------------------------------------
    # If Streamlit already knows that the user is authenticated,
    # still verify the Supabase user.
    # --------------------------------------------------------

    if st.session_state.get(
        "authenticated",
        False
    ):

        user = get_current_user()

        if user is None:

            st.session_state.user = None
            st.session_state.session = None
            st.session_state.profile = None
            st.session_state.is_admin = False
            st.session_state.authenticated = False

            return False


        # Refresh role/profile information.

        return refresh_authenticated_user()


    # --------------------------------------------------------
    # Check whether Supabase has an existing session.
    # --------------------------------------------------------

    session = get_current_session()

    if session is None:

        return False


    # --------------------------------------------------------
    # A session may exist but the user may no longer be valid.
    # --------------------------------------------------------

    user = get_current_user()

    if user is None:

        return False


    # --------------------------------------------------------
    # Store authenticated information.
    # --------------------------------------------------------

    return refresh_authenticated_user()


# ============================================================
#                    LOGIN RATE LIMIT
# ============================================================


def login_rate_limit_ok():

    """
    Basic application-side login throttling.

    Supabase Auth also provides its own protections.
    This adds a small additional layer against rapid repeated
    login attempts from the same Streamlit session.
    """

    now = time.time()

    last_time = float(
        st.session_state.get(
            "last_login_time",
            0.0
        )
    )

    attempts = int(
        st.session_state.get(
            "login_attempts",
            0
        )
    )


    # --------------------------------------------------------
    # Reset attempts after a quiet period.
    # --------------------------------------------------------

    if (
        last_time > 0
        and now - last_time > 300
    ):

        st.session_state.login_attempts = 0

        attempts = 0


    # --------------------------------------------------------
    # Allow up to 5 attempts in the current session window.
    # --------------------------------------------------------

    if attempts >= 5:

        if now - last_time < 300:

            return False


        st.session_state.login_attempts = 0

        return True


    return True


# ============================================================
#                    RECORD LOGIN ATTEMPT
# ============================================================


def record_login_attempt():

    st.session_state.login_attempts = (
        int(
            st.session_state.get(
                "login_attempts",
                0
            )
        )
        + 1
    )

    st.session_state.last_login_time = (
        time.time()
    )


# ============================================================
#                    LOGIN FORM
# ============================================================


def show_login_screen():

    """
    Display the login page.

    Users and administrators use the same Supabase
    authentication mechanism.

    Authorization is determined from the profiles table
    after successful authentication.
    """

    show_application_header()

    st.divider()

    # --------------------------------------------------------
    # Center-like layout
    # --------------------------------------------------------

    left, center, right = st.columns(
        [1, 2, 1]
    )


    with center:

        st.subheader(
            "Login"
        )

        st.write(
            "Sign in with your authorized account."
        )


        # ----------------------------------------------------
        # Existing login error
        # ----------------------------------------------------

        previous_error = (
            st.session_state.get(
                "last_error"
            )
        )


        if previous_error:

            st.error(
                previous_error
            )

            st.session_state.last_error = None


        # ----------------------------------------------------
        # Login form
        # ----------------------------------------------------

        with st.form(
            "login_form",
            clear_on_submit=False
        ):

            email = st.text_input(
                "Email",
                placeholder="Enter your email",
                autocomplete="email",
            )


            password = st.text_input(
                "Password",
                type="password",
                placeholder="Enter your password",
                autocomplete="current-password",
            )


            login_button = st.form_submit_button(
                "Login",
                use_container_width=True,
                type="primary",
            )


        # ----------------------------------------------------
        # Process login
        # ----------------------------------------------------

        if login_button:

            # -----------------------------------------------
            # Rate limit
            # -----------------------------------------------

            if not login_rate_limit_ok():

                st.error(
                    """
                    Too many login attempts.

                    Please wait a few minutes and try again.
                    """
                )

                return


            record_login_attempt()


            # -----------------------------------------------
            # Clean input
            # -----------------------------------------------

            email = clean_text(
                email,
                320
            ).lower()


            # -----------------------------------------------
            # Validate email before contacting Supabase
            # -----------------------------------------------

            if not is_valid_email(
                email
            ):

                st.error(
                    "Please enter a valid email address."
                )

                return


            if not password:

                st.error(
                    "Please enter your password."
                )

                return


            # -----------------------------------------------
            # Authenticate
            # -----------------------------------------------

            with st.spinner(
                "Signing in..."
            ):

                success, user, session, error = (
                    sign_in_user(
                        email,
                        password
                    )
                )


            # -----------------------------------------------
            # Login failed
            # -----------------------------------------------

            if not success:

                st.error(
                    error
                    or
                    "Invalid email or password."
                )

                return


            # -----------------------------------------------
            # Authentication succeeded.
            #
            # IMPORTANT:
            # Authentication does NOT automatically mean the
            # person is an authorized application user.
            #
            # The profile/role is checked next.
            # -----------------------------------------------

            st.session_state.user = user

            st.session_state.session = session

            st.session_state.authenticated = True


            # -----------------------------------------------
            # Verify application profile and role.
            # -----------------------------------------------

            authorized = (
                refresh_authenticated_user()
            )


            if not authorized:

                # refresh_authenticated_user() already clears
                # the authentication state when the account is
                # not authorized.

                st.error(
                    st.session_state.get(
                        "last_error"
                    )
                    or
                    "Your account is not authorized."
                )

                return


            # -----------------------------------------------
            # Successful login
            # -----------------------------------------------

            st.session_state.login_attempts = 0

            st.session_state.last_login_time = 0.0

            st.session_state.last_error = None


            # -----------------------------------------------
            # Rerun so the login page disappears immediately.
            # -----------------------------------------------

            st.rerun()


# ============================================================
#                    LOGOUT FUNCTION
# ============================================================


def logout_user():

    """
    Complete logout.

    This clears both the Supabase authentication session and
    the Streamlit-side application state.
    """

    try:

        if supabase is not None:

            supabase.auth.sign_out()

    except Exception:

        pass


    # --------------------------------------------------------
    # Clear all sensitive authentication-related state.
    # --------------------------------------------------------

    keys_to_clear = [
        "user",
        "session",
        "profile",
        "generated_output",
        "last_error",
    ]


    for key in keys_to_clear:

        if key in st.session_state:

            st.session_state[key] = None


    st.session_state.is_admin = False

    st.session_state.authenticated = False

    st.session_state.master_file_exists = False


    # --------------------------------------------------------
    # Clear login throttling information.
    # --------------------------------------------------------

    st.session_state.login_attempts = 0

    st.session_state.last_login_time = 0.0


# ============================================================
#                    SIDEBAR
# ============================================================


def show_sidebar():

    """
    Sidebar shown only after authentication.
    """

    user = st.session_state.get(
        "user"
    )

    if user is None:

        return


    email = get_user_email(
        user
    )


    is_admin = bool(
        st.session_state.get(
            "is_admin",
            False
        )
    )


    with st.sidebar:

        st.markdown(
            "## 6thSense"
        )

        st.caption(
            "6S-FO200 Vardaan"
        )

        st.divider()


        # ----------------------------------------------------
        # User information
        # ----------------------------------------------------

        st.write(
            f"**User:** {email}"
        )


        if is_admin:

            st.success(
                "ADMIN ACCESS"
            )

            st.write(
                "**Role:** Admin"
            )

        else:

            st.info(
                "USER ACCESS — Read-only"
            )

            st.write(
                "**Role:** User"
            )


        st.divider()


        # ----------------------------------------------------
        # Logout button
        # ----------------------------------------------------

        if st.button(
            "Logout",
            use_container_width=True,
            key="sidebar_logout_button",
        ):

            logout_user()

            st.rerun()


# ============================================================
#                    AUTHENTICATION GATE
# ============================================================
#
# Nothing below this section should be accessible until the
# user has passed authentication and application authorization.
#
# ============================================================


authenticated = initialize_authentication()


# ============================================================
#                    SHOW LOGIN OR APPLICATION
# ============================================================


if not authenticated:

    show_login_screen()

    st.stop()


# ============================================================
#                    RECHECK USER
# ============================================================


user = st.session_state.get(
    "user"
)


if user is None:

    st.session_state.authenticated = False

    show_login_screen()

    st.stop()


# ============================================================
#                    RECHECK USER PROFILE
# ============================================================
#
# Do not trust only the Streamlit session value.
# The role is refreshed from Supabase.
#
# ============================================================


user_id = get_user_id(
    user
)


if not user_id:

    logout_user()

    st.error(
        "Authentication could not be verified."
    )

    st.stop()


profile = get_user_profile(
    user_id
)


if profile is None:

    logout_user()

    st.error(
        "Your application authorization could not be verified."
    )

    st.stop()


# ============================================================
#                    ACTIVE ACCOUNT CHECK
# ============================================================


account_is_active = profile.get(
    "is_active",
    True
)


if not account_is_active:

    logout_user()

    st.error(
        "Your account has been disabled."
    )

    st.stop()


# ============================================================
#                    ROLE CHECK
# ============================================================


role = clean_text(
    profile.get(
        "role",
        USER_ROLE
    ),
    50
).lower()


if role not in {
    ADMIN_ROLE,
    USER_ROLE,
}:

    role = USER_ROLE


# ------------------------------------------------------------
# Store only the verified role.
# ------------------------------------------------------------

st.session_state.profile = profile

st.session_state.is_admin = (
    role == ADMIN_ROLE
)

st.session_state.authenticated = True


# ============================================================
#                    SIDEBAR
# ============================================================


show_sidebar()


# ============================================================
#                    LOGGED-IN USER
# ============================================================


user = st.session_state.user


st.title(
    "6thSense (6S-FO200) Vardaan"
)


st.caption(
    f"Logged in as: {get_user_email(user)}"
)


# ============================================================
#                    ACCESS STATUS
# ============================================================


if st.session_state.is_admin:

    st.success(
        "ADMIN ACCESS"
    )

else:

    st.info(
        "USER ACCESS — Read-only"
    )


# ============================================================
#                    END OF PART 2
# ============================================================
#
# >>> PART 3 STARTS DIRECTLY HERE <<<
#
# Part 3 will contain:
#
#   - Admin-only controls
#   - Master Excel upload
#   - Replace master file
#   - Delete master file
#   - Supabase Storage operations
#   - Master file metadata
#   - User management controls
#
# IMPORTANT:
# Users will NOT receive any admin controls.
#
# ============================================================
#
# >>> PART 2 STARTS DIRECTLY HERE <<<
#
# Part 2 will contain:
#
#   - Login screen
#   - Login form
#   - Authentication flow
#   - Logout function
#   - Sidebar
#   - Admin/User role display
#   - Session protection
#
# ============================================================
# ============================================================
# 6thSense (6S-FO200) Vardaan
# STREAMLIT + SUPABASE SECURE APPLICATION
#
# PART 3 OF 5
#
# Contains:
#   - Supabase Storage helpers
#   - Master Excel upload
#   - Replace master Excel
#   - Delete master Excel
#   - Master file metadata
#   - Admin-only user management
#   - Create user
#   - Enable / disable user
#   - Change user role
#
# CONTINUES DIRECTLY FROM PART 2
# ============================================================


# ============================================================
#                    STORAGE CONSTANTS
# ============================================================

STORAGE_BUCKET = "master-files"

MASTER_STORAGE_PATH = "master/master.xlsx"


# ============================================================
#                    ADMIN SECURITY CHECK
# ============================================================


def verify_admin_access():

    """
    Perform a fresh role check against Supabase.

    IMPORTANT:
    Streamlit session_state is never treated as the final
    authority for administrative operations.
    """

    user = st.session_state.get(
        "user"
    )

    if user is None:

        return False


    user_id = get_user_id(
        user
    )

    if not user_id:

        return False


    # Fresh database role check.

    profile = get_user_profile(
        user_id
    )


    if profile is None:

        return False


    if not profile.get(
        "is_active",
        True
    ):

        return False


    role = clean_text(
        profile.get(
            "role",
            ""
        ),
        50
    ).lower()


    return role == ADMIN_ROLE


# ============================================================
#                    STORAGE BUCKET CHECK
# ============================================================


def ensure_storage_bucket():

    """
    Check that the expected Supabase Storage bucket exists.

    Bucket creation should normally be done from Supabase
    Dashboard/SQL configuration rather than allowing arbitrary
    bucket creation from the application.
    """

    require_supabase()

    try:

        buckets = supabase.storage.list_buckets()

        if buckets is None:

            return False


        for bucket in buckets:

            # Supabase Python clients can return dictionaries
            # or objects depending on the installed version.

            if isinstance(
                bucket,
                dict
            ):

                bucket_name = bucket.get(
                    "name",
                    ""
                )

            else:

                bucket_name = getattr(
                    bucket,
                    "name",
                    ""
                )


            if bucket_name == STORAGE_BUCKET:

                return True


        return False


    except Exception:

        return False


# ============================================================
#                    DOWNLOAD MASTER FILE
# ============================================================


def download_master_file():

    """
    Download the currently stored master workbook from
    Supabase Storage.

    Returns:
        bytes or None
    """

    require_supabase()


    try:

        file_bytes = (
            supabase
            .storage
            .from_(STORAGE_BUCKET)
            .download(
                MASTER_STORAGE_PATH
            )
        )


        if not file_bytes:

            return None


        return file_bytes


    except Exception:

        return None


# ============================================================
#                    CHECK MASTER FILE
# ============================================================


def master_file_available():

    """
    Check whether the master workbook exists in Storage.
    """

    file_bytes = download_master_file()

    return file_bytes is not None


# ============================================================
#                    UPLOAD / REPLACE MASTER FILE
# ============================================================


def upload_master_file(
    file_bytes
):

    """
    Replace the existing master workbook.

    This function MUST only be called after verify_admin_access()
    succeeds.
    """

    require_supabase()


    if not verify_admin_access():

        return (
            False,
            "Administrative authorization failed."
        )


    if not file_bytes:

        return (
            False,
            "The uploaded file is empty."
        )


    # --------------------------------------------------------
    # Validate workbook before uploading.
    # --------------------------------------------------------

    try:

        test_buffer = io.BytesIO(
            file_bytes
        )

        workbook = load_workbook(
            test_buffer,
            read_only=True,
            data_only=False
        )

        workbook.close()

    except Exception:

        return (
            False,
            "The selected file is not a valid Excel workbook."
        )


    # --------------------------------------------------------
    # Calculate hash for metadata/logging.
    # --------------------------------------------------------

    file_hash = calculate_file_hash(
        file_bytes
    )


    try:

        response = (
            supabase
            .storage
            .from_(STORAGE_BUCKET)
            .upload(
                MASTER_STORAGE_PATH,
                file_bytes,
                {
                    "content-type":
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",

                    "upsert":
                        "true",
                }
            )
        )


        # Some Supabase versions return a response object and
        # some return a dictionary. We don't rely on its exact
        # shape here.
        _ = response


        # ----------------------------------------------------
        # Store file metadata when the optional table exists.
        # Failure of metadata logging should not make a
        # successfully uploaded file unusable.
        # ----------------------------------------------------

        try:

            admin_user = st.session_state.get(
                "user"
            )

            admin_id = get_user_id(
                admin_user
            )

            admin_email = get_user_email(
                admin_user
            )


            (
                supabase
                .table("master_file_metadata")
                .upsert(
                    {
                        "id": 1,
                        "file_name": MASTER_FILE_NAME,
                        "file_size": len(file_bytes),
                        "file_hash": file_hash,
                        "updated_by": admin_id,
                        "updated_by_email": admin_email,
                        "updated_at": utc_now(),
                    }
                )
                .execute()
            )

        except Exception:

            pass


        # ----------------------------------------------------
        # Update local state.
        # ----------------------------------------------------

        st.session_state.master_file_exists = True

        st.session_state.master_file_name = (
            MASTER_FILE_NAME
        )


        return (
            True,
            "Master file uploaded successfully."
        )


    except Exception:

        return (
            False,
            "The master file could not be uploaded."
        )


# ============================================================
#                    DELETE MASTER FILE
# ============================================================


def delete_master_file():

    """
    Delete the master workbook.

    ADMIN ONLY.
    """

    require_supabase()


    if not verify_admin_access():

        return (
            False,
            "Administrative authorization failed."
        )


    try:

        (
            supabase
            .storage
            .from_(STORAGE_BUCKET)
            .remove(
                [
                    MASTER_STORAGE_PATH
                ]
            )
        )


        # ----------------------------------------------------
        # Remove metadata.
        # ----------------------------------------------------

        try:

            (
                supabase
                .table("master_file_metadata")
                .delete()
                .eq(
                    "id",
                    1
                )
                .execute()
            )

        except Exception:

            pass


        # ----------------------------------------------------
        # Clear local state.
        # ----------------------------------------------------

        st.session_state.master_file_exists = False

        st.session_state.generated_output = None


        return (
            True,
            "Master file deleted successfully."
        )


    except Exception:

        return (
            False,
            "The master file could not be deleted."
        )


# ============================================================
#                    MASTER FILE INFORMATION
# ============================================================


def get_master_file_metadata():

    """
    Read metadata about the current master workbook.
    """

    require_supabase()


    try:

        response = (
            supabase
            .table("master_file_metadata")
            .select(
                "file_name,file_size,file_hash,updated_by_email,updated_at"
            )
            .eq(
                "id",
                1
            )
            .limit(1)
            .execute()
        )


        data = getattr(
            response,
            "data",
            None
        )


        if not data:

            return None


        return data[0]


    except Exception:

        return None


# ============================================================
#                    REFRESH MASTER STATUS
# ============================================================


def refresh_master_status():

    exists = master_file_available()

    st.session_state.master_file_exists = (
        exists
    )

    return exists


# ============================================================
#                    USER MANAGEMENT HELPERS
# ============================================================
#
# IMPORTANT:
#
# Supabase Auth user creation/deletion normally requires the
# service_role key.
#
# The service_role key must NEVER be placed in Streamlit
# frontend/session state or exposed to normal users.
#
# Therefore this application uses the safer architecture:
#
#   Supabase Auth -> authentication
#   profiles      -> application role/status
#
# User administration below manages application profiles.
#
# Creating the actual Auth account can be done through the
# Supabase Dashboard or through a secure server-side admin
# mechanism.
#
# We intentionally do NOT expose the service_role key here.
#
# ============================================================


def get_all_profiles():

    """
    Retrieve application users.

    ADMIN ONLY.
    """

    require_supabase()


    if not verify_admin_access():

        return []


    try:

        response = (
            supabase
            .table("profiles")
            .select(
                "id,email,role,is_active"
            )
            .order(
                "email"
            )
            .execute()
        )


        data = getattr(
            response,
            "data",
            None
        )


        if not data:

            return []


        return data


    except Exception:

        return []


# ============================================================
#                    UPDATE USER ROLE
# ============================================================


def update_user_role(
    user_id,
    new_role
):

    """
    Change application role.

    ADMIN ONLY.
    """

    require_supabase()


    if not verify_admin_access():

        return (
            False,
            "Administrative authorization failed."
        )


    user_id = clean_text(
        user_id,
        100
    )


    new_role = clean_text(
        new_role,
        50
    ).lower()


    if not user_id:

        return (
            False,
            "Invalid user."
        )


    if new_role not in {
        ADMIN_ROLE,
        USER_ROLE,
    }:

        return (
            False,
            "Invalid role."
        )


    current_admin_id = get_user_id()


    # --------------------------------------------------------
    # Prevent accidental removal of your own admin role.
    # --------------------------------------------------------

    if (
        user_id
        ==
        current_admin_id
        and
        new_role != ADMIN_ROLE
    ):

        return (
            False,
            "You cannot remove your own admin role."
        )


    try:

        (
            supabase
            .table("profiles")
            .update(
                {
                    "role": new_role
                }
            )
            .eq(
                "id",
                user_id
            )
            .execute()
        )


        return (
            True,
            "User role updated."
        )


    except Exception:

        return (
            False,
            "User role could not be updated."
        )


# ============================================================
#                    ENABLE / DISABLE USER
# ============================================================


def update_user_status(
    user_id,
    active
):

    """
    Enable or disable an application's user profile.

    ADMIN ONLY.
    """

    require_supabase()


    if not verify_admin_access():

        return (
            False,
            "Administrative authorization failed."
        )


    user_id = clean_text(
        user_id,
        100
    )


    if not user_id:

        return (
            False,
            "Invalid user."
        )


    current_admin_id = get_user_id()


    # --------------------------------------------------------
    # Prevent the currently logged-in administrator from
    # disabling their own account.
    # --------------------------------------------------------

    if (
        user_id
        ==
        current_admin_id
        and
        not bool(active)
    ):

        return (
            False,
            "You cannot disable your own account."
        )


    try:

        (
            supabase
            .table("profiles")
            .update(
                {
                    "is_active": bool(active)
                }
            )
            .eq(
                "id",
                user_id
            )
            .execute()
        )


        return (
            True,
            "User status updated."
        )


    except Exception:

        return (
            False,
            "User status could not be updated."
        )


# ============================================================
#                    ADMIN MASTER FILE PANEL
# ============================================================


def show_master_file_admin_panel():

    """
    Display master workbook controls.

    This function is called only for verified administrators.
    """

    # --------------------------------------------------------
    # SECOND authorization check.
    # --------------------------------------------------------

    if not verify_admin_access():

        st.error(
            "Administrator authorization failed."
        )

        return


    st.header(
        "Master File Management"
    )


    st.caption(
        "Only administrators can upload, replace, or delete the master file."
    )


    # --------------------------------------------------------
    # Current file status
    # --------------------------------------------------------

    exists = refresh_master_status()


    if exists:

        st.success(
            "Master file is available."
        )

    else:

        st.warning(
            "No master file is currently available."
        )


    # --------------------------------------------------------
    # File uploader
    # --------------------------------------------------------

    uploaded_file = st.file_uploader(
        "Upload / Replace Master Excel File",
        type=[
            "xlsx",
            "xlsm"
        ],
        key="admin_master_file_uploader",
    )


    if uploaded_file is not None:

        valid, error_message = (
            validate_excel_upload(
                uploaded_file
            )
        )


        if not valid:

            st.error(
                error_message
            )

        else:

            # ------------------------------------------------
            # Show file details before the destructive replace.
            # ------------------------------------------------

            st.write(
                f"**Selected file:** {uploaded_file.name}"
            )

            st.write(
                f"**Size:** {uploaded_file.size:,} bytes"
            )


            replace_confirmed = st.checkbox(
                "I confirm that this file should replace the current master file.",
                key="confirm_master_replace",
            )


            if st.button(
                "Upload / Replace Master File",
                type="primary",
                use_container_width=True,
                key="replace_master_button",
            ):

                if not replace_confirmed:

                    st.warning(
                        "Please confirm the replacement first."
                    )

                else:

                    # ----------------------------------------
                    # Read bytes once.
                    # ----------------------------------------

                    file_bytes = (
                        uploaded_file.getvalue()
                    )


                    # ----------------------------------------
                    # Final authorization immediately before
                    # write operation.
                    # ----------------------------------------

                    if not verify_admin_access():

                        st.error(
                            "Administrator authorization failed."
                        )

                    else:

                        with st.spinner(
                            "Uploading master file..."
                        ):

                            success, message = (
                                upload_master_file(
                                    file_bytes
                                )
                            )


                        if success:

                            st.success(
                                message
                            )

                            # Clear confirmation/uploader state
                            # after successful replacement.

                            st.session_state.confirm_master_replace = False

                            st.session_state.generated_output = None

                            st.rerun()

                        else:

                            st.error(
                                message
                            )


        # Delete current master file
    # --------------------------------------------------------

    st.divider()

    st.subheader(
        "Delete Master File"
    )


    if exists:

        delete_confirmed = st.checkbox(
            "I understand that deleting the master file will remove the current source workbook.",
            key="confirm_delete_master"
        )


        if st.button(
            "Delete Current Master File",
            key="delete_master_file",
            use_container_width=True
        ):

            if not delete_confirmed:

                st.warning(
                    "Please confirm that you want to delete the current master file."
                )

            else:

                try:

                    # ------------------------------------------------
                    # Security check — Admin only
                    # ------------------------------------------------

                    if not verify_admin_access():

                        st.error(
                            "Administrator access is required."
                        )

                    else:

                        with st.spinner(
                            "Deleting current master file..."
                        ):

                            supabase.storage \
                                .from_(STORAGE_BUCKET) \
                                .remove(
                                    [
                                        MASTER_STORAGE_PATH
                                    ]
                                )


                            # ----------------------------------------
                            # Delete metadata record
                            # ----------------------------------------

                            try:

                                (
                                    supabase
                                    .table(
                                        "master_file_metadata"
                                    )
                                    .delete()
                                    .eq(
                                        "storage_path",
                                        MASTER_STORAGE_PATH
                                    )
                                    .execute()
                                )

                            except Exception:

                                # Metadata deletion failure should
                                # not hide successful storage deletion.
                                pass


                        # --------------------------------------------
                        # Clear application state
                        # --------------------------------------------

                        st.session_state.master_file_exists = False

                        st.session_state.master_file_name = None

                        st.session_state.generated_output = None

                        st.session_state.generated_output_name = None


                        st.success(
                            "Current master file deleted successfully."
                        )


                        time.sleep(
                            0.8
                        )

                        st.rerun()


                except Exception as e:

                    st.error(
                        "Unable to delete the master file."
                    )

                    st.caption(
                        str(e)
                    )


    else:

        st.info(
            "No master file is currently available to delete."
        )
    # ============================================================
#              MASTER FILE METADATA
# ============================================================

def get_master_file_metadata():

    try:

        response = (
            supabase
            .table(
                "master_file_metadata"
            )
            .select(
                "*"
            )
            .eq(
                "storage_path",
                MASTER_STORAGE_PATH
            )
            .limit(
                1
            )
            .execute()
        )


        if response.data:

            return response.data[0]


        return None


    except Exception:

        return None
    # ============================================================
#              REFRESH MASTER FILE STATUS
# ============================================================

def refresh_master_status():

    try:

        exists = master_file_available()


        st.session_state.master_file_exists = (
            exists
        )


        if exists:

            st.session_state.master_file_name = (
                MASTER_FILE_NAME
            )

        else:

            st.session_state.master_file_name = None


        return exists


    except Exception:

        st.session_state.master_file_exists = False

        st.session_state.master_file_name = None

        return False
    # ============================================================
#                    USER MANAGEMENT
# ============================================================

def get_all_profiles():

    # --------------------------------------------------------
    # Security check
    # --------------------------------------------------------

    if not verify_admin_access():

        return []


    try:

        response = (
            supabase
            .table(
                "profiles"
            )
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

        show_safe_error(
            "Unable to load users.",
            e
        )

        return []
    # ============================================================
#                  UPDATE USER ROLE
# ============================================================

def update_user_role(
    user_id,
    new_role
):

    # --------------------------------------------------------
    # Security check — Admin only
    # --------------------------------------------------------

    if not verify_admin_access():

        return (
            False,
            "Administrator access is required."
        )


    # --------------------------------------------------------
    # Validate requested role
    # --------------------------------------------------------

    if new_role not in ROLES:

        return (
            False,
            "Invalid user role."
        )


    current_user_id = get_user_id()


    # --------------------------------------------------------
    # Safety:
    # Admin cannot remove their own administrator role
    # --------------------------------------------------------

    if str(user_id) == str(current_user_id):

        if new_role != "admin":

            return (
                False,
                "You cannot remove your own admin role."
            )


    try:

        response = (
            supabase
            .table(
                "profiles"
            )
            .update(
                {
                    "role": new_role
                }
            )
            .eq(
                "id",
                user_id
            )
            .execute()
        )


        return (
            True,
            "User role updated successfully."
        )


    except Exception as e:

        return (
            False,
            f"Unable to update user role: {e}"
    )
# ============================================================
#              UPDATE USER STATUS
# ============================================================

def update_user_status(
    user_id,
    active
):

    # --------------------------------------------------------
    # Security check — Admin only
    # --------------------------------------------------------

    if not verify_admin_access():

        return (
            False,
            "Administrator access is required."
        )


    current_user_id = get_user_id()


    # --------------------------------------------------------
    # Safety:
    # Admin cannot disable their own account
    # --------------------------------------------------------

    if str(user_id) == str(current_user_id):

        if not active:

            return (
                False,
                "You cannot disable your own account."
            )


    try:

        response = (
            supabase
            .table(
                "profiles"
            )
            .update(
                {
                    "is_active": bool(active)
                }
            )
            .eq(
                "id",
                user_id
            )
            .execute()
        )


        return (
            True,
            "User status updated successfully."
        )


    except Exception as e:

        return (
            False,
            f"Unable to update user status: {e}"
        )
# ============================================================
#              ADMIN USER MANAGEMENT PANEL
# ============================================================

def show_user_management_panel():

    # --------------------------------------------------------
    # Security check
    # --------------------------------------------------------

    if not verify_admin_access():

        st.error(
            "Administrator access is required."
        )

        return


    st.subheader(
        "User Management"
    )

    st.caption(
        "Admin can change user roles and enable or disable "
        "application access."
    )


    # --------------------------------------------------------
    # Load users
    # --------------------------------------------------------

    profiles = get_all_profiles()


    if not profiles:

        st.info(
            "No users found."
        )

        return


    current_user_id = get_user_id()


    # --------------------------------------------------------
    # Display users
    # --------------------------------------------------------

    for profile in profiles:

        user_id = profile.get(
            "id"
        )

        email = profile.get(
            "email",
            "Unknown"
        )

        current_role = profile.get(
            "role",
            "user"
        )

        current_status = profile.get(
            "is_active",
            True
        )

        created_at = profile.get(
            "created_at",
            ""
        )


        with st.container(
            border=True
        ):

            st.markdown(
                f"### {email}"
            )


            if str(user_id) == str(
                current_user_id
            ):

                st.caption(
                    "Current logged-in administrator"
                )


            # ------------------------------------------------
            # User information
            # ------------------------------------------------

            info_col1, info_col2 = st.columns(
                2
            )


            with info_col1:

                st.write(
                    f"**Current Role:** "
                    f"{current_role}"
                )


            with info_col2:

                status_text = (
                    "Active"
                    if current_status
                    else "Disabled"
                )

                st.write(
                    f"**Current Status:** "
                    f"{status_text}"
                )


            if created_at:

                st.caption(
                    f"Account created: {created_at}"
                )


            st.divider()


            # ------------------------------------------------
            # Controls
            # ------------------------------------------------

            control_col1, control_col2 = st.columns(
                2
            )


            with control_col1:

                selected_role = st.selectbox(
                    "Role",
                    options=[
                        "user",
                        "admin"
                    ],
                    index=(
                        1
                        if current_role == "admin"
                        else 0
                    ),
                    key=f"user_role_{user_id}"
                )


            with control_col2:

                selected_status = st.selectbox(
                    "Account Status",
                    options=[
                        "Active",
                        "Disabled"
                    ],
                    index=(
                        0
                        if current_status
                        else 1
                    ),
                    key=f"user_status_{user_id}"
                )


            # ------------------------------------------------
            # Save buttons
            # ------------------------------------------------

            save_col1, save_col2 = st.columns(
                2
            )


            with save_col1:

                if st.button(
                    "Save Role",
                    key=f"save_role_{user_id}",
                    use_container_width=True
                ):

                    success, message = (
                        update_user_role(
                            user_id,
                            selected_role
                        )
                    )


                    if success:

                        st.success(
                            message
                        )

                        time.sleep(
                            0.5
                        )

                        st.rerun()

                    else:

                        st.error(
                            message
                        )


            with save_col2:

                if st.button(
                    "Save Status",
                    key=f"save_status_{user_id}",
                    use_container_width=True
                ):

                    active_value = (
                        selected_status == "Active"
                    )


                    success, message = (
                        update_user_status(
                            user_id,
                            active_value
                        )
                    )


                    if success:

                        st.success(
                            message
                        )

                        time.sleep(
                            0.5
                        )

                        st.rerun()

                    else:

                        st.error(
                            message
                        )


            st.divider()
        # ============================================================
#                 ADMIN CONTROL NAVIGATION
# ============================================================

if st.session_state.is_admin:

    st.divider()

    st.header(
        "Administrator Controls"
    )

    # --------------------------------------------------------
    # Fresh administrator verification
    # --------------------------------------------------------

    admin_verified = verify_admin_access()


    if not admin_verified:

        st.warning(
            "Administrator privileges could not be verified. "
            "Please log in again."
        )

    else:

        admin_section = st.radio(
            "Select Administration Section",
            options=[
                "Master File",
                "User Management"
            ],
            horizontal=True,
            key="admin_navigation"
        )


        # ====================================================
        #                  MASTER FILE
        # ====================================================

        if admin_section == "Master File":

            show_master_file_admin_panel()


        # ====================================================
        #                  USER MANAGEMENT
        # ====================================================

        elif admin_section == "User Management":

            show_user_management_panel()
        # ============================================================
#              PART 4 — MASTER FILE PROCESSING
# ============================================================


# ============================================================
#              DOWNLOAD CURRENT MASTER FILE
# ============================================================

def download_current_master_file():

    # --------------------------------------------------------
    # Verify logged-in user
    # --------------------------------------------------------

    if not st.session_state.authenticated:

        return (
            None,
            "You must be logged in."
        )


    try:

        # ----------------------------------------------------
        # Check that master file exists
        # ----------------------------------------------------

        if not master_file_available():

            return (
                None,
                "No master file is currently available."
            )


        # ----------------------------------------------------
        # Download from private Supabase Storage
        # ----------------------------------------------------

        file_bytes = (
            supabase
            .storage
            .from_(STORAGE_BUCKET)
            .download(
                MASTER_STORAGE_PATH
            )
        )


        if not file_bytes:

            return (
                None,
                "The master file could not be downloaded."
            )


        # ----------------------------------------------------
        # Validate downloaded workbook
        # ----------------------------------------------------

        try:

            workbook = load_workbook(
                io.BytesIO(file_bytes),
                read_only=True,
                data_only=False
            )

            workbook.close()

        except Exception:

            return (
                None,
                "The stored master file is not a valid "
                "Excel workbook."
            )


        return (
            file_bytes,
            None
        )


    except Exception as e:

        return (
            None,
            f"Unable to download master file: {e}"
        )
    # ============================================================
#              LOAD MASTER WORKBOOK
# ============================================================

def load_current_master_workbook():

    # --------------------------------------------------------
    # Download master workbook
    # --------------------------------------------------------

    file_bytes, error_message = (
        download_current_master_file()
    )


    if error_message:

        return (
            None,
            None,
            error_message
        )


    if not file_bytes:

        return (
            None,
            None,
            "Master file is empty."
        )


    # --------------------------------------------------------
    # Load workbook into memory
    # --------------------------------------------------------

    try:

        workbook = load_workbook(
            io.BytesIO(
                file_bytes
            ),
            read_only=False,
            data_only=False
        )


    except Exception as e:

        return (
            None,
            None,
            f"Unable to open master workbook: {e}"
        )


    # --------------------------------------------------------
    # Safety check — workbook must contain worksheets
    # --------------------------------------------------------

    sheet_names = (
        workbook.sheetnames
    )


    if not sheet_names:

        workbook.close()

        return (
            None,
            None,
            "Master workbook does not contain any worksheets."
        )


    # --------------------------------------------------------
    # Return workbook + original bytes
    # --------------------------------------------------------

    return (
        workbook,
        file_bytes,
        None
    )
    # ============================================================
#              MASTER WORKBOOK SHEET INFORMATION
# ============================================================

def get_master_sheet_names():

    workbook = None


    try:

        workbook, _, error_message = (
            load_current_master_workbook()
        )


        if error_message:

            return []


        sheet_names = list(
            workbook.sheetnames
        )


        workbook.close()


        return sheet_names


    except Exception:

        if workbook is not None:

            try:

                workbook.close()

            except Exception:

                pass


        return []
    # ============================================================
#              EXCEL PROCESSING FUNCTION
# ============================================================

def process_master_workbook(
    master_bytes
):

    workbook = None

    try:

        # ----------------------------------------------------
        # Validate input
        # ----------------------------------------------------

        if not master_bytes:

            raise ValueError(
                "Master workbook data is empty."
            )


        # ----------------------------------------------------
        # Open master workbook
        # ----------------------------------------------------

        workbook = load_workbook(
            io.BytesIO(
                master_bytes
            ),
            read_only=False,
            data_only=False
        )


        # ----------------------------------------------------
        # Safety check
        # ----------------------------------------------------

        if not workbook.sheetnames:

            raise ValueError(
                "Master workbook contains no worksheets."
            )


        # ====================================================
        # IMPORTANT
        # ====================================================
        #
        # YOUR ORIGINAL WORKING EXCEL-PROCESSING CODE
        # MUST RUN HERE.
        #
        # Do NOT replace your existing calculations,
        # formulas, formatting, sorting, sheets, or output
        # logic with newly invented logic.
        #
        # The existing workbook-processing code should use:
        #
        #     workbook
        #
        # as its input workbook.
        #
        # ====================================================


        # ----------------------------------------------------
        # Temporary safety return
        # ----------------------------------------------------
        #
        # This prevents the app from pretending that an output
        # has been generated before the original processing
        # logic is connected.
        # ----------------------------------------------------

        raise NotImplementedError(
            "Original Excel-processing logic has not yet "
            "been connected."
        )


    except NotImplementedError:

        raise


    except Exception as e:

        raise RuntimeError(
            f"Excel processing failed: {e}"
        )


    finally:

        if workbook is not None:

            try:

                workbook.close()

            except Exception:

                pass
        # ============================================================
#              GENERATE OUTPUT WORKBOOK
# ============================================================

def generate_output_workbook():

    # --------------------------------------------------------
    # Authentication check
    # --------------------------------------------------------

    if not st.session_state.authenticated:

        return (
            None,
            "You must be logged in."
        )


    # --------------------------------------------------------
    # Active-account check
    # --------------------------------------------------------

    try:

        profile = get_user_profile(
            get_user_id()
        )


        if not profile:

            return (
                None,
                "User profile could not be verified."
            )


        if not profile.get(
            "is_active",
            True
        ):

            return (
                None,
                "Your account is disabled."
            )


    except Exception as e:

        return (
            None,
            f"Unable to verify account: {e}"
        )


    # --------------------------------------------------------
    # Load current master workbook
    # --------------------------------------------------------

    master_bytes, error_message = (
        download_current_master_file()
    )


    if error_message:

        return (
            None,
            error_message
        )


    if not master_bytes:

        return (
            None,
            "No master workbook is available."
        )


    # --------------------------------------------------------
    # Run the original Excel-processing logic
    # --------------------------------------------------------

    try:

        output_bytes = process_master_workbook(
            master_bytes
        )


        # ----------------------------------------------------
        # Validate processing result
        # ----------------------------------------------------

        if not output_bytes:

            return (
                None,
                "The processing function did not return "
                "an output workbook."
            )


        if not isinstance(
            output_bytes,
            (
                bytes,
                bytearray
            )
        ):

            return (
                None,
                "The processing function returned an "
                "invalid output format."
            )


        # ----------------------------------------------------
        # Validate generated Excel file
        # ----------------------------------------------------

        try:

            test_workbook = load_workbook(
                io.BytesIO(
                    output_bytes
                ),
                read_only=True,
                data_only=False
            )


            if not test_workbook.sheetnames:

                test_workbook.close()

                return (
                    None,
                    "Generated workbook contains no worksheets."
                )


            test_workbook.close()


        except Exception:

            return (
                None,
                "The generated output is not a valid "
                "Excel workbook."
            )


        return (
            bytes(output_bytes),
            None
        )


    except NotImplementedError:

        return (
            None,
            "The original Excel-processing code has not "
            "yet been connected to the application."
        )


    except Exception as e:

        return (
            None,
            f"Unable to generate output: {e}"
        )
# ============================================================
#                  GENERATE OUTPUT PANEL
# ============================================================

def show_generate_panel():

    st.divider()

    st.header(
        "Generate Output"
    )


    # --------------------------------------------------------
    # Current master status
    # --------------------------------------------------------

    exists = refresh_master_status()


    if not exists:

        st.warning(
            "No master file is currently available."
        )

        st.info(
            "Please ask the administrator to upload "
            "the master Excel file."
        )

        return


    st.success(
        f"Master file available: {MASTER_FILE_NAME}"
    )


    st.caption(
        "The current master workbook will be processed "
        "using the application's Excel-processing logic."
    )


    # --------------------------------------------------------
    # Generate button
    # --------------------------------------------------------

    if st.button(
        "Generate Output",
        key="generate_output_button",
        use_container_width=True,
        type="primary"
    ):

        # Clear previous output

        st.session_state.generated_output = None

        st.session_state.generated_output_name = None

        st.session_state.last_error = None


        with st.spinner(
            "Generating output workbook..."
        ):

            output_bytes, error_message = (
                generate_output_workbook()
            )


        if error_message:

            st.session_state.last_error = (
                error_message
            )

            st.error(
                error_message
            )


        else:

            st.session_state.generated_output = (
                output_bytes
            )

            st.session_state.generated_output_name = (
                OUTPUT_FILE_NAME
            )

            st.session_state.last_error = None


            st.success(
                "Output workbook generated successfully."
            )


    # --------------------------------------------------------
    # Download generated output
    # --------------------------------------------------------

    if st.session_state.generated_output:

        st.divider()

        st.subheader(
            "Generated File"
        )


        st.download_button(
            label="Download Output Excel",
            data=st.session_state.generated_output,
            file_name=(
                st.session_state.generated_output_name
                or OUTPUT_FILE_NAME
            ),
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            key="download_generated_output",
            use_container_width=True
        )
    # ============================================================
#                    USER DASHBOARD
# ============================================================


def show_user_dashboard():

    # --------------------------------------------------------
    # Authentication check
    # --------------------------------------------------------

    if not st.session_state.authenticated:

        st.error(
            "You must be logged in."
        )

        return


    # --------------------------------------------------------
    # Fresh account verification
    # --------------------------------------------------------

    user_id = get_user_id()

    if not user_id:

        st.error(
            "Unable to identify the logged-in user."
        )

        return


    profile = get_user_profile(
        user_id
    )


    if not profile:

        st.error(
            "Unable to verify your user profile."
        )

        return


    # --------------------------------------------------------
    # Active account check
    # --------------------------------------------------------

    if not profile.get(
        "is_active",
        True
    ):

        st.error(
            "Your account is disabled."
        )

        st.session_state.authenticated = False
        st.session_state.user = None
        st.session_state.session = None
        st.session_state.is_admin = False
        st.session_state.profile = None

        time.sleep(
            0.5
        )

        st.rerun()

        return


    # --------------------------------------------------------
    # User Dashboard
    # --------------------------------------------------------

    st.header(
        "User Dashboard"
    )

    st.caption(
        "View the current master file, generate the "
        "latest output workbook, and download it."
    )


    # --------------------------------------------------------
    # Master file status
    # --------------------------------------------------------

    exists = refresh_master_status()


    col1, col2 = st.columns(
        2
    )


    with col1:

        if exists:

            st.success(
                "Master File Available"
            )

            st.caption(
                f"Current file: {MASTER_FILE_NAME}"
            )

        else:

            st.warning(
                "Master File Not Available"
            )

            st.caption(
                "Please contact the administrator."
            )


    # --------------------------------------------------------
    # Refresh status
    # --------------------------------------------------------

    with col2:

        if st.button(
            "Refresh Master Status",
            key="user_refresh_master_status",
            use_container_width=True
        ):

            refresh_master_status()

            st.rerun()


    # --------------------------------------------------------
    # Generate / Download
    # --------------------------------------------------------

    show_generate_panel()
# ============================================================
#                  NORMAL USER APPLICATION
# ============================================================


if (
    st.session_state.authenticated
    and
    not st.session_state.is_admin
):

    show_user_dashboard()
# ============================================================
#                    ADMIN DASHBOARD
# ============================================================


def show_admin_dashboard():

    # --------------------------------------------------------
    # Authentication check
    # --------------------------------------------------------

    if not st.session_state.authenticated:

        st.error(
            "You must be logged in."
        )

        return


    # --------------------------------------------------------
    # Fresh administrator verification
    # --------------------------------------------------------

    if not verify_admin_access():

        st.error(
            "Administrator access could not be verified."
        )

        return


    # --------------------------------------------------------
    # Admin Dashboard
    # --------------------------------------------------------

    st.header(
        "Administrator Dashboard"
    )

    st.success(
        "ADMIN ACCESS"
    )

    st.caption(
        "You have full administrator access to the "
        "master workbook and user management."
    )


    # --------------------------------------------------------
    # Current master status
    # --------------------------------------------------------

    exists = refresh_master_status()


    if exists:

        st.success(
            f"Current Master File: {MASTER_FILE_NAME}"
        )

    else:

        st.warning(
            "No master file is currently available."
        )


    # --------------------------------------------------------
    # Normal Generate / Download functionality
    # --------------------------------------------------------

    show_generate_panel()


    # --------------------------------------------------------
    # Administrator controls
    # --------------------------------------------------------

    st.divider()

    st.header(
        "Administrator Controls"
    )


    admin_section = st.radio(
        "Select Administration Section",
        options=[
            "Master File",
            "User Management"
        ],
        horizontal=True,
        key="admin_dashboard_navigation"
    )


    if admin_section == "Master File":

        show_master_file_admin_panel()


    elif admin_section == "User Management":

        show_user_management_panel()
    # ============================================================
#                  MAIN DASHBOARD ROUTING
# ============================================================


if st.session_state.authenticated:

    if st.session_state.is_admin:

        show_admin_dashboard()

    else:

        show_user_dashboard()
    # ============================================================
#                 FINAL SESSION SAFETY CHECK
# ============================================================


def final_session_safety_check():

    # --------------------------------------------------------
    # User must be authenticated
    # --------------------------------------------------------

    if not st.session_state.get(
        "authenticated",
        False
    ):

        return False


    # --------------------------------------------------------
    # Get current user
    # --------------------------------------------------------

    user_id = get_user_id()

    if not user_id:

        st.session_state.authenticated = False
        st.session_state.user = None
        st.session_state.session = None
        st.session_state.is_admin = False
        st.session_state.profile = None

        return False


    # --------------------------------------------------------
    # Fresh profile verification
    # --------------------------------------------------------

    try:

        profile = get_user_profile(
            user_id
        )

    except Exception:

        st.session_state.authenticated = False
        st.session_state.user = None
        st.session_state.session = None
        st.session_state.is_admin = False
        st.session_state.profile = None

        return False


    if not profile:

        st.session_state.authenticated = False
        st.session_state.user = None
        st.session_state.session = None
        st.session_state.is_admin = False
        st.session_state.profile = None

        return False


    # --------------------------------------------------------
    # Active-account verification
    # --------------------------------------------------------

    if not profile.get(
        "is_active",
        True
    ):

        st.session_state.authenticated = False
        st.session_state.user = None
        st.session_state.session = None
        st.session_state.is_admin = False
        st.session_state.profile = None

        st.error(
            "Your account is currently disabled."
        )

        return False


    # --------------------------------------------------------
    # Refresh role from database
    # --------------------------------------------------------

    role = str(
        profile.get(
            "role",
            "user"
        )
    ).lower().strip()


    if role not in ROLES:

        role = "user"


    st.session_state.profile = profile

    st.session_state.is_admin = (
        role == "admin"
    )


    return True
# ============================================================
#              EXECUTE FINAL SESSION CHECK
# ============================================================


if st.session_state.authenticated:

    if not final_session_safety_check():

        st.warning(
            "Your session is no longer valid. "
            "Please log in again."
        )

        time.sleep(
            0.5
        )

        st.rerun()
    # ============================================================
#                  FINAL APPLICATION EXECUTION
# ============================================================


# ------------------------------------------------------------
# Final authentication/session validation
# ------------------------------------------------------------

if st.session_state.get(
    "authenticated",
    False
):

    if not final_session_safety_check():

        st.stop()


# ------------------------------------------------------------
# Refresh master status for logged-in users
# ------------------------------------------------------------

if st.session_state.get(
    "authenticated",
    False
):

    try:

        refresh_master_status()

    except Exception:

        st.session_state.master_file_exists = False
        st.session_state.master_file_name = None


# ------------------------------------------------------------
# Main application
# ------------------------------------------------------------

if st.session_state.get(
    "authenticated",
    False
):

    # --------------------------------------------------------
    # Administrator
    # --------------------------------------------------------

    if st.session_state.get(
        "is_admin",
        False
    ):

        show_admin_dashboard()


    # --------------------------------------------------------
    # Normal User
    # --------------------------------------------------------

    else:

        show_user_dashboard()
    # ============================================================
#                  LOGOUT & SESSION CLEANUP
# ============================================================


def clear_local_session_state():

    # --------------------------------------------------------
    # Clear authentication information
    # --------------------------------------------------------

    st.session_state.user = None

    st.session_state.session = None

    st.session_state.profile = None

    st.session_state.is_admin = False

    st.session_state.authenticated = False


    # --------------------------------------------------------
    # Clear generated workbook
    # --------------------------------------------------------

    st.session_state.generated_output = None

    st.session_state.generated_output_name = None


    # --------------------------------------------------------
    # Clear master-file status
    # --------------------------------------------------------

    st.session_state.master_file_exists = False

    st.session_state.master_file_name = None


    # --------------------------------------------------------
    # Clear temporary error information
    # --------------------------------------------------------

    st.session_state.last_error = None


def logout_user():

    try:

        # ----------------------------------------------------
        # Sign out from Supabase Auth
        # ----------------------------------------------------

        supabase.auth.sign_out()

    except Exception:

        # ----------------------------------------------------
        # Always clear local session even if Supabase
        # reports an error during logout.
        # ----------------------------------------------------

        pass


    # --------------------------------------------------------
    # Clear Streamlit session
    # --------------------------------------------------------

    clear_local_session_state()


    # --------------------------------------------------------
    # Return to login screen
    # --------------------------------------------------------

    st.rerun()


# ============================================================
#                    SIDEBAR LOGOUT
# ============================================================


with st.sidebar:

    st.divider()

    if st.button(
        "Logout",
        key="sidebar_logout_button",
        use_container_width=True
    ):

        logout_user()
    # ============================================================
#             PART 5 — EXCEL PROCESSING CONNECTION
# ============================================================


def process_master_workbook(
    master_bytes
):

    workbook = None

    try:

        # --------------------------------------------------------
        # Validate input
        # --------------------------------------------------------

        if not master_bytes:

            raise ValueError(
                "Master workbook data is empty."
            )


        # --------------------------------------------------------
        # Open master workbook
        # --------------------------------------------------------

        workbook = load_workbook(
            io.BytesIO(
                master_bytes
            ),
            read_only=False,
            data_only=False
        )


        # --------------------------------------------------------
        # Validate worksheets
        # --------------------------------------------------------

        if not workbook.sheetnames:

            raise ValueError(
                "Master workbook contains no worksheets."
            )


        # ========================================================
        # YOUR ORIGINAL WORKING EXCEL CODE GOES HERE
        # ========================================================
        #
        # IMPORTANT:
        #
        # Do NOT change your existing Excel calculations.
        #
        # Your original code should:
        #
        # 1. Read the master workbook
        # 2. Perform all existing calculations
        # 3. Preserve all existing sheets/output
        # 4. Add the requested columns/calculations
        # 5. Create the final summary_output.xlsx
        # 6. Save the resulting workbook into memory
        #
        # The final result MUST be returned as bytes.
        #
        # ========================================================


        raise NotImplementedError(
            "Your original working Excel-processing code "
            "must be inserted here."
        )


    except NotImplementedError:

        raise


    except Exception as e:

        raise RuntimeError(
            f"Excel processing failed: {e}"
        )


    finally:

        if workbook is not None:

            try:

                workbook.close()

            except Exception:

                pass
        # ============================================================
#        PART 5 — ORIGINAL EXCEL PROCESSOR CONNECTION
# ============================================================


def process_master_workbook(master_bytes):

    if not master_bytes:

        raise ValueError(
            "Master workbook is empty."
        )


    # --------------------------------------------------------
    # Open uploaded/current master workbook
    # --------------------------------------------------------

    input_buffer = io.BytesIO(
        master_bytes
    )


    try:

        source_workbook = load_workbook(
            input_buffer,
            read_only=False,
            data_only=False
        )

    except Exception as e:

        raise RuntimeError(
            f"Unable to open master workbook: {e}"
        )


    try:

        if not source_workbook.sheetnames:

            raise ValueError(
                "Master workbook contains no worksheets."
            )


        # ====================================================
        # ORIGINAL WORKING EXCEL LOGIC
        # ====================================================
        #
        # IMPORTANT:
        #
        # Your existing working Excel-processing code must
        # be placed here without changing its calculations.
        #
        # It must ultimately create the final workbook in
        # memory and return its bytes.
        #
        # Existing output requirements such as:
        #
        # - existing worksheets
        # - existing formulas
        # - existing columns
        # - existing calculations
        # - sorting
        # - formatting
        # - filters
        # - Sum O2H.10
        # - Sum O2L.10
        # - Vol.Expand columns
        # - all other existing output
        #
        # must remain unchanged.
        #
        # ====================================================


        raise NotImplementedError(
            "Original working Excel-processing code "
            "has not yet been connected."
        )


    finally:

        try:

            source_workbook.close()

        except Exception:

            pass
        ```sql
-- ============================================================
--        6thSense (6S-FO200) Vardaan
--        SUPABASE DATABASE SECURITY / RLS
-- ============================================================


-- ============================================================
-- 1. PROFILES TABLE
-- ============================================================

create table if not exists public.profiles (

    id uuid primary key references auth.users(id) on delete cascade,

    email text,

    role text not null default 'user'
        check (role in ('admin', 'user')),

    is_active boolean not null default true,

    created_at timestamptz not null default now()

);


-- ============================================================
-- 2. MASTER FILE METADATA TABLE
-- ============================================================

create table if not exists public.master_file_metadata (

    id bigint generated by default as identity primary key,

    file_name text not null,

    storage_path text not null unique,

    file_size bigint,

    file_hash text,

    uploaded_by uuid references auth.users(id),

    uploaded_at timestamptz not null default now()

);


-- ============================================================
-- 3. ENABLE ROW LEVEL SECURITY
-- ============================================================

alter table public.profiles
enable row level security;


alter table public.master_file_metadata
enable row level security;


-- ============================================================
-- 4. HELPER FUNCTION
--    Check whether logged-in user is an ADMIN
-- ============================================================

create or replace function public.is_admin()
returns boolean
language sql
stable
security definer
set search_path = public
as $$

    select exists (

        select 1

        from public.profiles

        where id = auth.uid()

        and role = 'admin'

        and is_active = true

    );

$$;


-- ============================================================
-- 5. PROFILES POLICIES
-- ============================================================


-- Users can read their own profile.

drop policy if exists
"Users can read own profile"
on public.profiles;


create policy
"Users can read own profile"

on public.profiles

for select

to authenticated

using (

    id = auth.uid()

    or public.is_admin()

);


-- Admin can update user roles/status.

drop policy if exists
"Admins can update profiles"
on public.profiles;


create policy
"Admins can update profiles"

on public.profiles

for update

to authenticated

using (

    public.is_admin()

)

with check (

    public.is_admin()

);


-- ============================================================
-- 6. MASTER FILE METADATA POLICIES
-- ============================================================


-- Logged-in users can see metadata.

drop policy if exists
"Authenticated users can read master metadata"
on public.master_file_metadata;


create policy
"Authenticated users can read master metadata"

on public.master_file_metadata

for select

to authenticated

using (

    true

);


-- Only administrators can insert metadata.

drop policy if exists
"Admins can insert master metadata"
on public.master_file_metadata;


create policy
"Admins can insert master metadata"

on public.master_file_metadata

for insert

to authenticated

with check (

    public.is_admin()

);


-- Only administrators can update metadata.

drop policy if exists
"Admins can update master metadata"
on public.master_file_metadata;


create policy
"Admins can update master metadata"

on public.master_file_metadata

for update

to authenticated

using (

    public.is_admin()

)

with check (

    public.is_admin()

);


-- Only administrators can delete metadata.

drop policy if exists
"Admins can delete master metadata"
on public.master_file_metadata;


create policy
"Admins can delete master metadata"

on public.master_file_metadata

for delete

to authenticated

using (

    public.is_admin()

);


-- ============================================================
-- 7. IMPORTANT
-- ============================================================
--
-- DO NOT PUT THE SUPABASE SERVICE_ROLE KEY
-- INSIDE STREAMLIT SECRETS.
--
-- The Streamlit application must use only the
-- normal Supabase client key.
--
-- ============================================================
```sql
-- ============================================================
--        6thSense (6S-FO200) Vardaan
--        SUPABASE STORAGE SECURITY
-- ============================================================


-- ============================================================
-- 1. CREATE STORAGE BUCKET
-- ============================================================
--
-- The bucket is PRIVATE.
-- Users will NOT get a public URL to the master file.
--

insert into storage.buckets
(
    id,
    name,
    public
)

values
(
    'master-files',
    'master-files',
    false
)

on conflict (id)
do update set
    public = false;


-- ============================================================
-- 2. STORAGE RLS
-- ============================================================

-- Logged-in users can DOWNLOAD the master file.
--
-- This is intentionally limited to authenticated users.
--

drop policy if exists
"Authenticated users can download master file"
on storage.objects;


create policy
"Authenticated users can download master file"

on storage.objects

for select

to authenticated

using
(
    bucket_id = 'master-files'
);


-- ============================================================
-- 3. ONLY ADMIN CAN UPLOAD
-- ============================================================

drop policy if exists
"Admins can upload master file"
on storage.objects;


create policy
"Admins can upload master file"

on storage.objects

for insert

to authenticated

with check
(
    bucket_id = 'master-files'

    and public.is_admin()
);


-- ============================================================
-- 4. ONLY ADMIN CAN UPDATE / REPLACE
-- ============================================================

drop policy if exists
"Admins can replace master file"
on storage.objects;


create policy
"Admins can replace master file"

on storage.objects

for update

to authenticated

using
(
    bucket_id = 'master-files'

    and public.is_admin()
)

with check
(
    bucket_id = 'master-files'

    and public.is_admin()
);


-- ============================================================
-- 5. ONLY ADMIN CAN DELETE
-- ============================================================

drop policy if exists
"Admins can delete master file"
on storage.objects;


create policy
"Admins can delete master file"

on storage.objects

for delete

to authenticated

using
(
    bucket_id = 'master-files'

    and public.is_admin()
);


-- ============================================================
-- 6. IMPORTANT SECURITY NOTE
-- ============================================================
--
-- The bucket MUST remain PRIVATE.
--
-- Normal users:
--     ✓ Can download through the authenticated application
--     ✓ Can view master-file status
--     ✓ Can generate output
--     ✓ Can download generated output
--     ✗ Cannot upload
--     ✗ Cannot replace
--     ✗ Cannot delete
--
-- Administrators:
--     ✓ Upload
--     ✓ Replace
--     ✓ Delete
--     ✓ Manage users
--
-- NEVER make the master-files bucket public.
--
-- NEVER put the Supabase service_role key in:
--     Streamlit secrets
--     GitHub
--     app.py
--     requirements.txt
--     browser/client code
--
-- ============================================================



    
