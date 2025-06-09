import streamlit as st 
import pandas as pd
from xlsx2ldf import ExcelToLDFConverter
import os
from datetime import datetime
import re

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


def main():
    st.markdown('<h1 class="title">📄 Excel to LDF Converter</h1>', unsafe_allow_html=True)
    st.markdown("Upload your Excel file containing LIN data to convert it to an LDF file.")

    col1, col2 = st.columns([3, 1])

    with col1:
        uploaded_file = st.file_uploader("Choose an Excel file", type=["xls", "xlsx"], key="file_uploader")

    with col2:
        if st.button("Convert to LDF", key="convert_button"):
            with st.spinner('Converting to LDF... Please wait'):
                converter = ExcelToLDFConverter(uploaded_file.name)
                if converter.convert("custom_filename.ldf"):
                            st.success("Conversion completed successfully!")

                            with open("custom_filename.ldf", "rb") as f:
                                bytes_data = f.read()
                                st.download_button(
                                    label="Download LDF File",
                                    data=bytes_data,
                                    file_name="custom_filename.ldf",
                                    mime="application/octet-stream",
                                    key="download_button"
                                )
                else:
                    st.error("Conversion failed. Please check the input data.")
            
if __name__ == "__main__":
    main()