import pandas as pd
from streamlit.runtime.uploaded_file_manager import UploadedFile
from typing import List, Union, Dict
import re
import pprint
import streamlit as st
import os
import math

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
    .error-box {
        background-color: #ffebee;
        border-left: 5px solid #f44336;
        padding: 10px;
        margin: 10px 0;
        border-radius: 5px;
    }
    .warning-box {
        background-color: #fff8e1;
        border-left: 5px solid #ffc107;
        padding: 10px;
        margin: 10px 0;
        border-radius: 5px;
    }
    .success-box {
        background-color: #e8f5e9;
        border-left: 5px solid #4caf50;
        padding: 10px;
        margin: 10px 0;
        border-radius: 5px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

def load_xlsx(file_path: str) -> Union[pd.DataFrame, Dict]:
    try:
        if isinstance(file_path, str) or isinstance(file_path, UploadedFile):
            data_frame = pd.read_excel(
                file_path, sheet_name="ETH.Matrix", keep_default_na=True, engine="openpyxl"
            )
            return data_frame
        elif isinstance(file_path, List):
            finally_df = {}
            for file in file_path:
                data_frame = pd.read_excel(
                    file, sheet_name="ETH.Matrix", keep_default_na=True, engine="openpyxl"
                )
                if isinstance(file, UploadedFile):
                    finally_df[file.name] = data_frame
                else:
                    finally_df[file.split("\\")[-1]] = data_frame
            return finally_df
    except Exception as e:
        return f"Undefined type of file: {e}"


def create_correct_df(df: pd.DataFrame) -> pd.DataFrame:

    new_df_data = {
        "Rename Topic": df["Rename Topic"].ffill(),
        # "IAT Recommend Topic": df["IAT Recommend Topic"].ffill(),
        "Cloud pub": df["Cloud"].ffill(),
        "Cloud broker": df["Unnamed: 3"].ffill(),
        "Cloud sub": df["Unnamed: 4"].ffill(),
        "Vehicle pub": df["In-Vehicle"].ffill(),
        "Vehicle broker": df["Unnamed: 6"].ffill(),
        "Vehicle sub": df["Unnamed: 7"].ffill(),
        "Topic content type": df["Topic Content Type"].ffill(),
        "API Description": df["Vehicle API description"],
        "ETH signal name": df["ETH Signal Name"],
        "DBC Msg name": df["DBC message name"],
        "DBC Sig name": df["DBC signal name"],
        "Datadesription": df["Datadescription"],
        "Unit": df["Unit"],
        "Datatype": df["Datatype"],
        "Init val": df["Initial Value"],
        "Min": df["Min Value"],
        "Max": df["Max Value"],
        "Coding val": df["CodingValue-Enum"],
        "Comments": df["Comments"]

    }

    new_df = pd.DataFrame(new_df_data)

    new_df["Unit"] = new_df["Unit"].astype(str)
    new_df["Unit"] = new_df["Unit"].str.replace("Ω", "Ohm", regex=False)
    new_df["Unit"] = new_df["Unit"].str.replace("℃", "degC", regex=False)

    new_df = new_df.dropna(subset=["DBC Sig name"])

    agg_funcs = {
        col: 'first' for col in new_df.columns
    }

    new_df = new_df.groupby("Rename Topic", as_index=False).agg(agg_funcs)

    return new_df



def main():
    st.title("🚧ETH Topics Validator")
    uploaded_file = st.file_uploader("Upload matrix file", type=["xlsx"])

    if uploaded_file:
        try:
            df = load_xlsx(uploaded_file)
            processed_df = create_correct_df(df)

            st.success("File loaded successfully!")

            st.dataframe(processed_df)
           
        except Exception as e:
            st.error(e)


if __name__ == "__main__":
    main()