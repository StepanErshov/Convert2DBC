import streamlit as st
import pandas as pd
from xlsx2ldf import ExcelToLDFConverter
import os
from datetime import datetime
import re
import logging
from logging.handlers import RotatingFileHandler


def setup_logging():
    """
    Настройка журнала регистрации ошибок и действий.
    Создает журнал в папке logs с ротацией старых записей.
    """
    log_dir = "logs"
    os.makedirs(log_dir, exist_ok=True)
    
    log_format = "%(asctime)s - %(levelname)s - %(message)s"
    date_format = "%Y-%m-%d %H:%M:%S"
    
    logging.basicConfig(
        level=logging.INFO,
        format=log_format,
        datefmt=date_format,
        handlers=[
            RotatingFileHandler(
                filename=os.path.join(log_dir, "excel_to_ldf.log"),
                maxBytes=5 * 1024 * 1024,  # 5 MB
                backupCount=3
            )
        ]
    )

    console = logging.StreamHandler()
    console.setLevel(logging.INFO)
    formatter = logging.Formatter(log_format, date_format)
    console.setFormatter(formatter)
    logging.getLogger('').addHandler(console)


setup_logging()


st.set_page_config(
    page_title="Excel to LDF Converter",
    page_icon="🚗",
    layout="wide"
)

st.markdown("""
    <style>
    .main {
        background-color: #f5f5f5;
    }
    .stButton > button {
        background-color: #4CAF50;
        color: white;
        border-radius: 5px;
        padding: 10px 24px;
    }
    .stButton > button:hover {
        background-color: #45a049;
    }
    .stFileUploader > div > div > div > button {
        background-color: #2196F3;
        color: white;
    }
    .stTextInput > div > div > input {
        border-radius: 5px;
    }
    .title {
        color: #2c3e50;
    }
    </style>
""", unsafe_allow_html=True)


def extract_version_date(filename):
    """
    Извлекает номер версии и дату из имени файла.
    Возвращает кортеж (version, date).
    """
    pattern = r"_V(\d+\.\d+\.\d+)_(\d{8})\."
    match = re.search(pattern, filename)
    if match:
        return match.group(1), match.group(2)
    return None, None


def generate_base_name(input_filename):
    """
    Генерирует базовое имя файла без номера версии и даты.
    """
    base_name = os.path.splitext(input_filename)[0]
    base_name = re.sub(r"_V\d+\.\d+\.\d+_\d{8}$", "", base_name)
    return base_name


def generate_default_output_filename(input_filename, new_version=None):
    """
    Генерация стандартного имени выходного файла.
    """
    base_name = generate_base_name(input_filename)
    current_date = datetime.now().strftime("%Y%m%d")
    
    if new_version is None:
        version, _ = extract_version_date(input_filename)
        new_version = version if version else "1.0.0"
    
    return f"{base_name}_V{new_version}_{current_date}.ldf"


def main():
    """
    Основная логика интерфейса Streamlit.
    Обрабатывает загрузку файла, отображение данных и выполнение преобразования.
    """
    st.markdown('<h1 class="title">📄 Excel to LDF Converter</h1>', unsafe_allow_html=True)
    st.markdown("Загрузите Excel-файл, содержащий данные LIN, для преобразования в LDF-файл.")

    col1, col2 = st.columns([3, 1])

    with col1:
        uploaded_file = st.file_uploader("Выберите Excel-файл", type=["xls", "xlsx"], key="file_uploader")
        
        if uploaded_file is not None:
            try:
                df = pd.read_excel(uploaded_file, sheet_name="Matrix")
                st.subheader("Предпросмотр данных")
                st.dataframe(df.head().style.set_properties(**{
                    'background-color': '#f0f2f6',
                    'color': '#2c3e50',
                    'border': '1px solid #dfe6e9'
                }))
            except Exception as e:
                st.error(f"Произошла ошибка при чтении Excel-файла: {str(e)}")
                return

    with col2:
        if uploaded_file is not None:
            st.subheader("Настройки вывода")
            
            version, _ = extract_version_date(uploaded_file.name)
            default_version = version if version else "1.0.0"
            
            new_version = st.text_input(
                "Версия LDF",
                value=default_version,
                help="Укажите номер версии формата X.X.X"
            )
            
            base_name = generate_base_name(uploaded_file.name)
            
            default_output_name = generate_default_output_filename(uploaded_file.name, new_version)
            custom_filename = st.text_input(
                "Имя выходного LDF-файла",
                value=default_output_name,
                help="Вы можете изменить название выходного файла"
            )
            
            if not custom_filename.lower().endswith('.ldf'):
                custom_filename += '.ldf'
                
            st.markdown("**Итоговое имя LDF-файла:**")
            st.code(custom_filename)
            
            if st.button("Преобразовать в LDF"):
                with st.spinner('Выполняем преобразование в LDF... Подождите немного'):
                    try:
                        converter = ExcelToLDFConverter(uploaded_file)
                        
                        if converter.convert(custom_filename):
                            st.success("Преобразование успешно выполнено!")
                            
                            with open(custom_filename, "rb") as f:
                                bytes_data = f.read()
                                st.download_button(
                                    label="Скачать LDF-файл",
                                    data=bytes_data,
                                    file_name=custom_filename,
                                    mime="application/octet-stream",
                                    key="download_button"
                                )
                        else:
                            st.error("Ошибка при преобразовании. Пожалуйста, проверьте входные данные.")
                    except Exception as e:
                        st.error(f"Возникла ошибка при выполнении преобразования: {str(e)}")


if __name__ == "__main__":
    main()