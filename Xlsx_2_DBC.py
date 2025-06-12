import streamlit as st
import pandas as pd
from xlsx2dbc import ExcelToDBCConverter
import os
from datetime import datetime
import re

st.set_page_config(
    page_title="Excel to DBC Converter", page_icon=":car:", layout="wide"
)

st.markdown(
    """
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
    """,
    unsafe_allow_html=True,
)


def extract_version_date(filename):
    pattern = r"_V(\d+\.\d+\.\d+)_(\d{8})\."
    match = re.search(pattern, filename)
    if match:
        return match.group(1), match.group(2)
    return None, None


def generate_base_name(input_filename):
    base_name = os.path.splitext(input_filename)[0]
    base_name = re.sub(r"_V\d+\.\d+\.\d+_\d{8}$", "", base_name)
    return base_name


def generate_default_output_filename(input_filename, new_version=None):
    base_name = generate_base_name(input_filename)
    current_date = datetime.now().strftime("%Y%m%d")

    if new_version is None:
        version, _ = extract_version_date(input_filename)
        new_version = version if version else "1.0.0"

    return f"{base_name}_V{new_version}_{current_date}.dbc"


def main():
    st.markdown(
        '<h1 class="title">📊 Excel to DBC Converter</h1>', unsafe_allow_html=True
    )
    st.markdown(
        "Upload your Excel file containing CAN data to convert it to a DBC file."
    )

    col1, col2 = st.columns([3, 1])

    with col1:
        uploaded_file = st.file_uploader(
            "Choose an Excel file", type=["xlsx"], key="file_uploader"
        )

        if uploaded_file is not None:
            try:
                df = pd.read_excel(uploaded_file, sheet_name="Matrix")
                st.subheader("Data Preview")
                st.dataframe(
                    df.head().style.set_properties(
                        **{
                            "background-color": "#f0f2f6",
                            "color": "#2c3e50",
                            "border": "1px solid #dfe6e9",
                        }
                    )
                )
            except Exception as e:
                st.error(f"Error reading the Excel file: {str(e)}")
                return

    with col2:
        if uploaded_file is not None:
            st.subheader("Output Settings")

            version, _ = extract_version_date(uploaded_file.name)
            default_version = version if version else "1.0.0"

            new_version = st.text_input(
                "DBC Version",
                value=default_version,
                help="Enter the version number in format X.X.X",
            )

            base_name = generate_base_name(uploaded_file.name)

            default_output_name = generate_default_output_filename(
                uploaded_file.name, new_version
            )
            custom_filename = st.text_input(
                "Output DBC file name",
                value=default_output_name,
                help="You can customize the output file name",
            )

            if not custom_filename.lower().endswith(".dbc"):
                custom_filename += ".dbc"

            st.markdown("**Final DBC file name:**")
            st.code(custom_filename)

            if st.button("Convert to DBC", key="convert_button"):
                with st.spinner("Converting... Please wait"):
                    try:
                        converter = ExcelToDBCConverter(uploaded_file)
                        if converter.convert(custom_filename):
                            st.success("Conversion completed successfully!")

                            with open(custom_filename, "rb") as f:
                                bytes_data = f.read()
                                st.download_button(
                                    label="Download DBC File",
                                    data=bytes_data,
                                    file_name=custom_filename,
                                    mime="application/octet-stream",
                                    key="download_button",
                                )
                        else:
                            st.error("Conversion failed. Please check the input data.")
                    except Exception as e:
                        st.error(f"An error occurred during conversion: {str(e)}")


if __name__ == "__main__":
    main()
