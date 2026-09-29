import streamlit as st

import uuid
import sqlite3
from datetime import datetime

from utils.theme import load_css

from modules import (

    #AMAZON MODULES
    amazon_xml,
    amazon_aci_json,
    #amazon_xml_gets,

    #APC MODULES
    apc,
    apc_billing,
    apc_billing_header_report,
    apc_client_details,
    apc_pallet_id,
    apc_candata,
    apc_billing_validation,
    apc_hscode_validation,

    #DATA PROCESSING MODULES
    candata,
    aci_json,
    airshipment,
    prohibited,

    #UTILITY MODULES
    defender,
    ezclear,
    spliptpdf,
    compresspdf,
    hscode,

    #OTHER MODULES  
    #hr_payroll,
    #test_sharepoint,
    #excelmerger,
    #gst,
    #candata_to_gets_format,
    #bandofcanada,
    #amazon_xml_old,
    #vendor,
)


# --------PAGE CONFIG----------

st.set_page_config(
    page_title="UAT ABS Enterprise Portal",
    page_icon="assets/qwe1.ico",
    layout="wide",
    initial_sidebar_state="expanded"
)

def login():

    # Already authenticated
    if st.session_state.get("authenticated", False):
        return True

    # Hide sidebar while logged out
    st.markdown(
        """
        <style>
            [data-testid="stSidebar"] {
                display: none;
            }
        </style>
        """,
        unsafe_allow_html=True
    )

    st.markdown(
        """
        <div style="
            max-width: 450px;
            margin: 100px auto 20px auto;
            padding: 35px;
            border-radius: 15px;
            background: white;
            box-shadow: 0 4px 20px rgba(0,0,0,0.15);
            text-align: center;
        ">
            <h1>🔐 ACB Enterprise Portal</h1>
            <p style="color:#666;">
                Please sign in to continue
            </p>
        </div>
        """,
        unsafe_allow_html=True
    )

    username = st.text_input(
        "Username",
        key="login_username"
    )

    password = st.text_input(
        "Password",
        type="password",
        key="login_password"
    )

    if st.button(
        "🔐 Login",
        use_container_width=True
    ):

        # Get all users from secrets
        users = st.secrets["credentials"]["users"]

        # Check username and password
        if username in users and password == users[username]:

            st.session_state.authenticated = True
            st.session_state.username = username

            st.session_state.pop("login_username", None)
            st.session_state.pop("login_password", None)

            st.rerun()

        else:
            st.error("❌ Invalid username or password.")

    return False

if not login():
    st.stop()
# --------LOAD THEME----------
load_css()

st.sidebar.markdown(
    f"👤 **Logged in as:** {st.session_state.username}"
)

# -------------STATE-----------
if "module" not in st.session_state:
    st.session_state.module = "HOME"

if "active_group" not in st.session_state:
    st.session_state.active_group = None

if "session_id" not in st.session_state:
    st.session_state.session_id = str(uuid.uuid4())

# ---- SIDEBAR NAVIGATION (ENTERPRISE STYLE) ----------

NAV_GROUPS = {

    "🚚 APC": {
        "🚚 APC sFTP": "APC",
        "💳 APC BILLING DETAIL": "APC_BILLING",
        "💳 APC CLIENT DETAILS": "APC_CLIENT_DETAILS",
        "💳 APC BILLING HEADER": "APC_BILLING_HEADER",
        "📦 APC PALLET ID": "APC_PALLET_ID",
        "📊 APC CANDATA UPLOAD FILE": "APC_CANDATA",
        "💳 APC BILLING VALIDATION": "APC_BILLING_VALIDATION",
        "🔎 APC HSCODE VALIDATION": "APC_HSCODE_VALIDATION"
    },

    "📦 AMAZON": {
        "📊 AMAZON XML TO CANDATA": "AMAZON_XML",
        "📦 AMAZON ACI JSON": "AMAZON_ACI_JSON",
    },
    
    "📊 Data Processing": {
        "📊 CANDATA UPLOAD FILE": "CANDATA",
        "📦 ACI JSON": "ACI_JSON",
        "✈️ AIR SHIPMENT":"AIRSHIPMENT",
        "📄 PROHIBITED ITEM DETECTION":"PROHIBITED",
    },

    "🛠 Utilities": {
        "🛡️ DEFENDER": "DEFENDER",
        "📑 Split PDF": "SPLIT_PDF",
        "📦 Compress PDF": "compresspdf",
        "📊 EZCLEAR": "ezclear",
        "🔎 HSCODE SEARCH": "hscode",
        #"💰 HR PAYROLL": "hr_payroll",
    }
}

st.sidebar.markdown(
    '''
    <div class="sidebar-menu-title"> ☰ MENU</div>
    <div class="sidebar-menu-caption">NAVIGATION</div>
    ''',
    unsafe_allow_html=True
)

for group_name, items in NAV_GROUPS.items():

    is_open = st.session_state.active_group == group_name

    with st.sidebar.expander(group_name, expanded=is_open):

        for label, module in items.items():

            if st.button(
                label,
                key=f"{group_name}_{module}"
            ):

                st.session_state.module = module
                st.session_state.active_group = group_name
                st.rerun()

MODULES = {

    #AMAZON MODULES
    "AMAZON_XML": amazon_xml.run,
    "AMAZON_ACI_JSON": amazon_aci_json.run,
    #"AMAZON_XML_GETS": amazon_xml_gets.run,

    #DATA PROCESSING MODULES
    "CANDATA": candata.run,
    "ACI_JSON": aci_json.run,
    "AIRSHIPMENT": airshipment.run,
    "PROHIBITED": prohibited.run,

    #APC MODULES
    "APC": apc.run,
    "APC_BILLING": apc_billing.run,
    "APC_CLIENT_DETAILS": apc_client_details.run,
    "APC_BILLING_HEADER": apc_billing_header_report.run,
    "APC_PALLET_ID": apc_pallet_id.run,
    "APC_CANDATA": apc_candata.run,
    "APC_BILLING_VALIDATION": apc_billing_validation.run,
    "APC_HSCODE_VALIDATION": apc_hscode_validation.run,
    
    #UTILITY MODULES
    "DEFENDER": defender.run,
    "SPLIT_PDF": spliptpdf.run,
    "ezclear": ezclear.run,
    "compresspdf": compresspdf.run,
    "hscode": hscode.run,
    #"hr_payroll": hr_payroll.run,
    #"excelmerger": excelmerger.run,
    #"gst": gst.run,
    #"candata_to_gets_format": candata_to_gets_format.run,
}

# ---------------- HOME PAGE ----------------
if st.session_state.module == "HOME":
    
    st.markdown('<div class="main-header">🚀 ACB Enterprise Portal</div>', unsafe_allow_html=True)
    st.caption('<div class="sub-header">Centralized Automation & Validation System</div>', unsafe_allow_html=True)

    st.markdown("""
    <div class="card">
        Welcome to ACB Toolkit. Select a module from the sidebar to begin processing.
    </div>
    """, unsafe_allow_html=True)

    st.markdown("---")
    st.caption("© 2026 ACB Toolkit | Developed by IT Department")
else:

    module = st.session_state.module

    if module in MODULES:
        try:
            MODULES[module]()
        except Exception as e:
            st.error(f"Failed loading {module}")
            st.exception(e)

    else:
        st.error(f"Module {module} not registered")