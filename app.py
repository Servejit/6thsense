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
