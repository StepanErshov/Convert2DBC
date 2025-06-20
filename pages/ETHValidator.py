import pandas as pd
from streamlit.runtime.uploaded_file_manager import UploadedFile
from typing import List, Union, Dict
import re
import pprint
import streamlit as st
import os
import math

# st.markdown(
#     """
#     <style>
#     .main {
#         background-color: #f5f5f5;
#     }
#     .stButton>button {
#         background-color: #4CAF50;
#         color: white;
#         border-radius: 5px;
#         padding: 10px 24px;
#     }
#     .stButton>button:hover {
#         background-color: #45a049;
#     }
#     .stFileUploader>div>div>div>button {
#         background-color: #2196F3;
#         color: white;
#     }
#     .stTextInput>div>div>input {
#         border-radius: 5px;
#     }
#     .title {
#         color: #2c3e50;
#     }
#     .error-box {
#         background-color: #ffebee;
#         border-left: 5px solid #f44336;
#         padding: 10px;
#         margin: 10px 0;
#         border-radius: 5px;
#     }
#     .warning-box {
#         background-color: #fff8e1;
#         border-left: 5px solid #ffc107;
#         padding: 10px;
#         margin: 10px 0;
#         border-radius: 5px;
#     }
#     .success-box {
#         background-color: #e8f5e9;
#         border-left: 5px solid #4caf50;
#         padding: 10px;
#         margin: 10px 0;
#         border-radius: 5px;
#     }
#     </style>
#     """,
#     unsafe_allow_html=True,
# )

# def load_xlsx(file_path: str) -> Union[pd.DataFrame, Dict]:
#     try:
#         if isinstance(file_path, str) or isinstance(file_path, UploadedFile):
#             data_frame = pd.read_excel(
#                 file_path, sheet_name="ETH.Matrix", keep_default_na=True, engine="openpyxl"
#             )
#             return data_frame
#         elif isinstance(file_path, List):
#             finally_df = {}
#             for file in file_path:
#                 data_frame = pd.read_excel(
#                     file, sheet_name="ETH.Matrix", keep_default_na=True, engine="openpyxl"
#                 )
#                 if isinstance(file, UploadedFile):
#                     finally_df[file.name] = data_frame
#                 else:
#                     finally_df[file.split("\\")[-1]] = data_frame
#             return finally_df
#     except Exception as e:
#         return f"Undefined type of file: {e}"


# def create_correct_df(df: pd.DataFrame) -> pd.DataFrame:
#     bus_users = [
#         col
#         for col in df.columns
#         if any(val in ["S", "R"] for val in df[col].dropna().unique())
#         and col != "Unit\n单位"
#     ]
    
#     print(bus_users)

#     senders = []
#     receivers = []

#     for _, row in df.iterrows():
#         row_senders = []
#         row_receivers = []

#         for bus_user in bus_users:
#             if bus_user in df.columns:
#                 if pd.notna(row[bus_user]) and row[bus_user] == "S":
#                     row_senders.append(bus_user)
#                 elif pd.notna(row[bus_user]) and row[bus_user] == "R":
#                     row_receivers.append(bus_user)

#         senders.append(",".join(row_senders) if row_senders else "Vector__XXX")
#         receivers.append(",".join(row_receivers) if row_receivers else "Vector__XXX")

#     new_df_data = {
#         "Msg ID": df["Msg ID(hex)\n报文标识符"].ffill(),
#         "Msg Name": df["Msg Name\n报文名称"].ffill(),
#         "Protected ID": df["Protected ID (hex)\n保护标识符"].ffill(),
#         "Send Type": df["Msg Send Type\n报文发送类型"].ffill(),
#         "Checksum Mode": df["Checksum mode\n校验方式"].ffill(),
#         "Msg Length": df["Msg Length(Byte)\n报文长度"].ffill(),
#         "Sig Name": df["Signal Name\n信号名称"],
#         "Description": df["Signal Description\n信号描述"],
#         "Response Error": df["Response Error"],
#         "Start Byte": df["Start Byte\n起始字节"],
#         "Start Bit": df["Start Bit\n起始位"],
#         "Length": df["Bit Length(Bit)\n信号长度"],
#         "Resolution": df["Resolution\n精度"],
#         "Offset": df["Offset\n偏移量"],
#         "Min": df["Signal Min. Value(phys)\n物理最小值"],
#         "Max": df["Signal Max. Value(phys)\n物理最大值"],
#         "Min Hex": df["Signal Min. Value(Hex)\n总线最小值"],
#         "Max Hex": df["Signal Max. Value(Hex)\n总线最大值"],
#         "Unit": df["Unit\n单位"],
#         "Initinal": df["Initial Value(Hex)\n初始值"],
#         "Invalid": df["Invalid Value(Hex)\n无效值"],
#         "Signal Value Description": df["Signal Value Description(hex)\n信号值描述"],
#         "Remark": df["Remark\n备注"],
#         "Receiver": receivers,
#         "Senders": senders,
#     }

#     new_df = pd.DataFrame(new_df_data)

#     new_df["Unit"] = new_df["Unit"].astype(str)
#     new_df["Unit"] = new_df["Unit"].str.replace("Ω", "Ohm", regex=False)
#     new_df["Unit"] = new_df["Unit"].str.replace("℃", "degC", regex=False)

#     new_df = new_df.dropna(subset=["Sig Name"])
    
#     new_df["Is Signed"] = False

#     return bus_users





# pprint.pprint(create_correct_df(load_xlsx("C:\\projects\\Convert2DBC\\ATOM_Ethernet_Matrix_V4.1.3_20250220.xlsx")))