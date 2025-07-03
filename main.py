import streamlit as st
from streamlit_extras.app_logo import add_logo

# import dbc2xlsx


st.set_page_config(
    page_title="CAN/LIN Tools Suite",
    page_icon="🚗",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
<style>
    .stPageLink a {
        border-radius: 8px !important;
        padding: 10px 14px !important;
        margin: 6px 0 !important;
        transition: all 0.2s ease !important;
    }
    
    .stPageLink a:hover {
        background-color: #f0f0f0 !important;
    }
    
    .nav-section-header {
        font-size: 0.9rem;
        color: #555;
        margin-top: 15px;
        margin-bottom: 5px;
        text-transform: uppercase;
        font-weight: 600;
    }
    
    .logo {
        margin-bottom: 20px;
        text-align: center;
    }
</style>
""",
    unsafe_allow_html=True,
)

xlsx2dbc = st.Page("pages/Xlsx_2_DBC.py", title="Xlsx 2 DBC", icon="📊")

dbc2xlsx = st.Page("pages/DBC_2_Xlsx.py", title="DBC 2 Xlsx", icon="📊")

routing_tables = st.Page("pages/Routing_table.py", title="Routing Tables", icon="🔄")

domain2ecu = st.Page("pages/Domain_2_ECU.py", title="Domain 2 ECU", icon="⚙️")

xls2ldf = st.Page("pages/Xls_2_LDF.py", title="Xls 2 LDF", icon="📈")

can_validator = st.Page("pages/CANValidator.py", title="CAN Validator", icon="🚧")

lin_validator = st.Page("pages/LINValidator.py", title="LIN Validator", icon="🚀")

# eth_validator = st.Page(
#     "pages/ETHValidator.py",
#     title="ETH Validator",
#     icon="⚠️"
# )

pg = st.navigation(
    {
        "Main tools": [
            xlsx2dbc,
            routing_tables,
            domain2ecu,
            xls2ldf,
            can_validator,
            lin_validator,
        ],
        "Developments": [dbc2xlsx],
    }
)

pg.run()

if st.session_state.get("current_page") == "home":
    st.title("Welcome in CAN/LIN Tools Suite")
    st.markdown(
        """
    ### Available tools:
    - **Xlsx 2 DBC** - Excel to DBC format converter
    - **Routing tables** - Routing tables management
    - **Domain 2 ECU** - Domain and ECU mapping
    - **Xls 2 LDF** - Excel to LIN description format converter
    - **CAN Validator** - CAN matrix validator (under development)
    """
    )

    st.divider()
    st.info("Choose tool from navigation menu on the left")
