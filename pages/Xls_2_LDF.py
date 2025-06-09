import streamlit as st
import pandas as pd
from xlsx2ldf import ExcelToLDFConverter
import os
from datetime import datetime
import re
import tempfile

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
    .stButton>button {
        background-color: #4CAF50;
        color: white;
        border-radius: 5px;
        padding: 10px 24px;
    }
    .stButton>button:hover {
        background-color: #45a049;
    }
    .stFileUploader>div>div>div>button {
        background-color: #2196F3;
        color: white;
    }
    .stTextInput>div>div>input {
        border-radius: 5px;
    }
    .title {
        color: #2c3e50;
    }
    </style>
    """, unsafe_allow_html=True)

def extract_version_date(filename):
    pattern = r'_V(\d+\.\d+\.\d+)_(\d{8})\.'
    match = re.search(pattern, filename)
    if match:
        return match.group(1), match.group(2)
    return None, None

def generate_base_name(input_filename):
    base_name = os.path.splitext(input_filename)[0]
    base_name = re.sub(r'_V\d+\.\d+\.\d+_\d{8}$', '', base_name)
    return base_name

def generate_default_output_filename(input_filename, new_version=None):
    base_name = generate_base_name(input_filename)
    current_date = datetime.now().strftime("%Y%m%d")
    
    if new_version is None:
        version, _ = extract_version_date(input_filename)
        new_version = version if version else "1.0.0"
    
    return f"{base_name}_V{new_version}_{current_date}.ldf"

def save_uploaded_file(uploaded_file):
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            return tmp_file.name
    except Exception as e:
        st.error(f"Error saving temporary file: {str(e)}")
        return None

def main():
    st.markdown('<h1 class="title">📄 Excel to LDF Converter</h1>', unsafe_allow_html=True)
    st.markdown("Upload your Excel file containing LIN data to convert it to an LDF file.")

    col1, col2 = st.columns([3, 1])

    with col1:
        uploaded_file = st.file_uploader("Choose an Excel file", type=["xls", "xlsx"], key="file_uploader")

        if uploaded_file is not None:
            try:
                df = pd.read_excel(uploaded_file, sheet_name="Matrix")
                st.subheader("Data Preview")
                st.dataframe(df.head().style.set_properties(**{
                    'background-color': '#f0f2f6',
                    'color': '#2c3e50',
                    'border': '1px solid #dfe6e9'
                }))
            except Exception as e:
                st.error(f"Error reading the Excel file: {str(e)}")
                return

    with col2:
        if uploaded_file is not None:
            st.subheader("Output Settings")

            version, _ = extract_version_date(uploaded_file.name)
            default_version = version if version else "1.0.0"

            new_version = st.text_input(
                "LDF Version",
                value=default_version,
                help="Enter the version number in format X.X.X"
            )

            base_name = generate_base_name(uploaded_file.name)

            default_output_name = generate_default_output_filename(uploaded_file.name, new_version)
            custom_filename = st.text_input(
                "Output LDF file name",
                value=default_output_name,
                help="You can customize the output file name"
            )

            if not custom_filename.lower().endswith('.ldf'):
                custom_filename += '.ldf'

            st.markdown("**Final LDF file name:**")
            st.code(custom_filename)

            if st.button("Convert to LDF", key="convert_button"):
                with st.spinner('Converting to LDF... Please wait'):
                    try:
                        # Сохраняем временный файл
                        temp_file_path = save_uploaded_file(uploaded_file)
                        if not temp_file_path:
                            st.error("Failed to create temporary file")
                            return
                            
                        # Создаем конвертер и конвертируем
                        converter = ExcelToLDFConverter(temp_file_path)
                        
                        # Сохраняем результат во временный файл
                        with tempfile.NamedTemporaryFile(delete=False, suffix='.ldf') as output_tmp:
                            output_path = output_tmp.name
                        
                        if converter.convert(output_path):
                            st.success("Conversion completed successfully!")
                            
                            # Читаем результат и предлагаем скачать
                            with open(output_path, "rb") as f:
                                bytes_data = f.read()
                                st.download_button(
                                    label="Download LDF File",
                                    data=bytes_data,
                                    file_name=custom_filename,
                                    mime="application/octet-stream",
                                    key="download_button"
                                )
                        else:
                            st.error("Conversion failed. Please check the input data.")
                            
                        # Удаляем временные файлы
                        try:
                            os.unlink(temp_file_path)
                            os.unlink(output_path)
                        except:
                            pass
                            
                    except Exception as e:
                        st.error(f"An error occurred during conversion: {str(e)}")
                        # Удаляем временные файлы в случае ошибки
                        try:
                            if 'temp_file_path' in locals() and os.path.exists(temp_file_path):
                                os.unlink(temp_file_path)
                            if 'output_path' in locals() and os.path.exists(output_path):
                                os.unlink(output_path)
                        except:
                            pass

if __name__ == "__main__":
    main()