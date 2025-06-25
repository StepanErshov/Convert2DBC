import streamlit as st
from streamlit_extras.app_logo import add_logo
import streamlit_keycloak as keycloak
import os

# Инициализация Keycloak (секреты из переменных окружения)
keycloak = keycloak.login(
    url=os.getenv("KEYCLOAK_URL", "https://your-keycloak-domain.com/auth"),
    realm=os.getenv("KEYCLOAK_REALM", "cmdtool"),
    client_id=os.getenv("KEYCLOAK_CLIENT_ID", "streamlit-app"),
    client_secret_key=os.getenv("KEYCLOAK_CLIENT_SECRET"),
    init_options={
        "checkLoginIframe": False,
        "onLoad": "login-required"
    }
)

# Проверка аутентификации
if not keycloak.authenticated:
    st.warning("Пожалуйста, войдите для доступа к приложению")
    st.stop()

# Основная конфигурация приложения
st.set_page_config(
    page_title="CAN/LIN Tools Suite",
    page_icon="🚗",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Добавьте ваш логотип (убедитесь что путь правильный)
add_logo("assets/logo.png", height=80)

# Стили (оставьте ваши текущие стили)
st.markdown("""
<style>
    /* Ваши текущие стили */
</style>
""", unsafe_allow_html=True)

# Проверка ролей (пример)
if 'admin' not in keycloak.roles:
    st.error("У вас недостаточно прав для доступа")
    st.stop()

# Ваше текущее меню навигации
xlsx2dbc = st.Page("pages/Xlsx_2_DBC.py", title="Xlsx 2 DBC", icon="📊")
routing_tables = st.Page("pages/Routing_table.py", title="Routing Tables", icon="🔄")
domain2ecu = st.Page("pages/Domain_2_ECU.py", title="Domain 2 ECU", icon="⚙️")
xls2ldf = st.Page("pages/Xls_2_LDF.py", title="Xls 2 LDF", icon="📈")
can_validator = st.Page("pages/CANValidator.py", title="CAN Validator", icon="🚧")
lin_validator = st.Page("pages/LINValidator.py", title="LIN Validator", icon="⚠️")

pg = st.navigation({
    "Main tools": [xlsx2dbc, routing_tables, domain2ecu, xls2ldf, can_validator],
    "Developments": [lin_validator],
})

pg.run()

# Дополнительная информация о пользователе в sidebar
with st.sidebar:
    if keycloak.authenticated:
        st.write(f"Добро пожаловать, {keycloak.username}!")
        if st.button("Выйти"):
            keycloak.logout()