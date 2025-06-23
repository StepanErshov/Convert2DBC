import pandas as pd
from streamlit.runtime.uploaded_file_manager import UploadedFile
from typing import List, Union, Dict
import re
import pprint
import streamlit as st
import os
import math

# st.set_page_config(page_title="CAN Validator", page_icon="⚠️", layout="wide")

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


def get_file_info(file_name: str):
    file_start = "ATOM_CAN_Matrix_"
    file_start1 = "ATOM_CANFD_Matrix_"
    file_name_only = os.path.splitext(os.path.basename(file_name))[0]
    if file_name_only.startswith(file_start1):
        protocol = "CANFD"
        start_index = 0
        parts = file_name_only[len(file_start1) :].split("_")
    elif file_name_only.startswith(file_start):
        protocol = "CAN"
        start_index = 0
        parts = file_name_only[len(file_start) :].split("_")
    else:
        protocol = ""
    if not (
        file_name_only.startswith(file_start) or file_name_only.startswith(file_start1)
    ):
        return None
    start_index = file_name_only.find(file_start1)
    if start_index != -1:
        parts = file_name_only[start_index + len(file_start1) :].split("_")
    else:
        parts = file_name_only[len(file_start) :].split("_")
    domain_name = parts.pop(0)
    version_string = parts.pop(0)
    if version_string.startswith("V"):
        version = version_string[1:]
        versions = version.split(".")
        if len(versions) != 3:
            return None
    else:
        version = ""
    file_date = parts.pop(0)
    if len(parts) > 0:
        if parts[0] == "internal":
            parts.pop(0)
        device_name = "_".join(parts)
    else:
        device_name = ""

    return {
        "version": version,
        "date": file_date,
        "device_name": device_name,
        "domain_name": domain_name,
        "protocol": protocol,
    }


def load_xlsx(file_path: str) -> Union[pd.DataFrame, Dict]:
    try:
        if isinstance(file_path, str) or isinstance(file_path, UploadedFile):
            data_frame = pd.read_excel(
                file_path, sheet_name="Matrix", keep_default_na=True, engine="openpyxl"
            )
            return data_frame
        elif isinstance(file_path, List):
            finally_df = {}
            for file in file_path:
                data_frame = pd.read_excel(
                    file, sheet_name="Matrix", keep_default_na=True, engine="openpyxl"
                )
                if isinstance(file, UploadedFile):
                    finally_df[file.name] = data_frame
                else:
                    finally_df[file.split("\\")[-1]] = data_frame
            return finally_df
    except Exception as e:
        return f"Undefined type of file: {e}"


def create_correct_df(df: pd.DataFrame) -> pd.DataFrame:
    bus_users = [
        col
        for col in df.columns
        if any(val in ["S", "R"] for val in df[col].dropna().unique())
        and col != "Unit\n单位"
    ]
    senders = []
    receivers = []

    for _, row in df.iterrows():
        row_senders = []
        row_receivers = []

        for bus_user in bus_users:
            if bus_user in df.columns:
                if pd.notna(row[bus_user]) and row[bus_user] == "S":
                    row_senders.append(bus_user)
                elif pd.notna(row[bus_user]) and row[bus_user] == "R":
                    row_receivers.append(bus_user)

        senders.append(",".join(row_senders) if row_senders else "Vector__XXX")
        receivers.append(",".join(row_receivers) if row_receivers else "Vector__XXX")

    new_df_data = {
        "Msg ID": df["Msg ID\n报文标识符"].ffill(),
        "Msg Name": df["Msg Name\n报文名称"].ffill(),
        "Cycle Type": df["Msg Cycle Time (ms)\n报文周期时间"].ffill(),
        "Msg Time Fast": df["Msg Cycle Time Fast(ms)\n报文发送的快速周期"].ffill(),
        "Msg Reption": df["Msg Nr. Of Reption\n报文快速发送的次数"].ffill(),
        "Msg Delay": df["Msg Delay Time(ms)\n报文延时时间"].ffill(),
        "Msg Type": df["Msg Type\n报文类型"].ffill(),
        "Send Type": df["Msg Send Type\n报文发送类型"].ffill(),
        "Msg Length": df["Msg Length (Byte)\n报文长度"].ffill(),
        "Sig Name": df["Signal Name\n信号名称"],
        "Start Byte": df["Start Byte\n起始字节"],
        "Start Bit": df["Start Bit\n起始位"],
        "Length": df["Bit Length (Bit)\n信号长度"],
        "Resolution": df["Resolution\n精度"],
        "Offset": df["Offset\n偏移量"],
        "Initinal": df["Initial Value (Hex)\n初始值"],
        "Invalid": df["Invalid Value(Hex)\n无效值"],
        "Min": df["Signal Min. Value (phys)\n物理最小值"],
        "Min Hex": df["Signal Min. Value (Hex)\n总线最小值"],
        "Max": df["Signal Max. Value (phys)\n物理最大值"],
        "Max Hex": df["Signal Max. Value (Hex)\n总线最大值"],
        "Unit": df["Unit\n单位"],
        "Receiver": receivers,
        "Byte Order": df["Byte Order\n排列格式(Intel/Motorola)"],
        "Data Type": df["Data Type\n数据类型"],
        "Description": df["Signal Description\n信号描述"],
        "Signal Value Description": df["Signal Value Description\n信号值描述"],
        "Senders": senders,
        "Signal Send Type": df["Signal Send Type\n信号发送类型"],
        "Inactive value": df["Inactive Value (Hex)\n非使能值"],
    }

    if "BRS\n传输速率切换标识位" in df.columns:
        new_df_data["BRS"] = df["BRS\n传输速率切换标识位"].ffill()
    else:
        new_df_data["BRS"] = None

    if "Frame Format\n帧格式" in df.columns:
        new_df_data["Frame Format"] = df["Frame Format\n帧格式"].ffill()
    else:
        new_df_data["Frame Format"] = None

    new_df = pd.DataFrame(new_df_data)

    new_df["Unit"] = new_df["Unit"].astype(str)
    new_df["Unit"] = new_df["Unit"].str.replace("Ω", "Ohm", regex=False)
    new_df["Unit"] = new_df["Unit"].str.replace("℃", "degC", regex=False)

    new_df = new_df.dropna(subset=["Sig Name"])
    new_df["Is Signed"] = new_df["Data Type"].str.contains("Signed", na=False)

    return new_df


def export_validation_errors_to_excel(data_frame: pd.DataFrame, file_path: str) -> None:
    error_sheets = {}
    print(data_frame)

    # 1. Message Name errors
    invalid_names = []
    too_long_names = []
    msg_names = set(data_frame["Msg Name"].dropna().astype(str))
    for name in msg_names:
        if not re.fullmatch(r"^[A-Za-z0-9_\-]+$", name.strip()):
            invalid_names.append(name)
        if len(name) > 64:
            too_long_names.append(name)

    if invalid_names:
        error_sheets["Invalid_Message_Names"] = pd.DataFrame(
            {"Incorrect name": invalid_names}
        )
    if too_long_names:
        error_sheets["Long_Message_Names"] = pd.DataFrame(
            {"Name": too_long_names, "Length": [len(n) for n in too_long_names]}
        )

    # 2. Message Type errors
    msg_type = dict(zip(data_frame["Msg Name"], data_frame["Msg Type"]))
    invalid_type = {}
    invalid_name = {}
    for key, val in msg_type.items():
        if val not in ["Normal", "Diag", "NM"]:
            invalid_type[key] = val
        if key.startswith("Diag") and val != "Diag":
            invalid_name[key] = val
        if key.startswith("NM_") and val != "NM":
            invalid_name[key] = val

    if invalid_type:
        error_sheets["Invalid_Message_Types"] = pd.DataFrame(
            {
                "Message Name": invalid_type.keys(),
                "Incorrect Type": invalid_type.values(),
            }
        )
    if invalid_name:
        error_sheets["Message_Name_Type_Mismatch"] = pd.DataFrame(
            {"Message Name": invalid_name.keys(), "Current Type": invalid_name.values()}
        )

    # 3. Message ID errors
    data_frame["Msg ID"] = data_frame["Msg ID"].apply(
        lambda x: int(x, 16) if isinstance(x, str) and x.startswith("0x") else int(x)
    )
    msg_id = dict(zip(data_frame["Msg Name"], data_frame["Msg ID"]))
    msg_type = dict(zip(data_frame["Msg Name"], data_frame["Msg Type"]))
    invalid_id = {}
    invalid_type = {}
    for mes, id in msg_id.items():
        if not (0x001 <= id <= 0x7FF):
            invalid_id[mes] = id
        if 0x700 <= id <= 0x7FF and msg_type[mes] != "Diag":
            invalid_type[mes] = id
        if 0x500 <= id <= 0x5FF and msg_type[mes] != "NM":
            invalid_type[mes] = id

    if invalid_id:
        error_sheets["Invalid_Message_IDs"] = pd.DataFrame(
            {"Message Name": invalid_id.keys(), "Invalid ID": invalid_id.values()}
        )
    if invalid_type:
        error_sheets["Message_ID_Type_Mismatch"] = pd.DataFrame(
            {"Message Name": invalid_type.keys(), "Invalid ID": invalid_type.values()}
        )

    # 4. Message Send Type errors
    msg_send_type = dict(zip(data_frame["Msg Name"], data_frame["Send Type"]))
    invalid_send_type = {}
    for mes, send_type in msg_send_type.items():
        if send_type not in ["Cycle", "Event", "CE"]:
            invalid_send_type[mes] = send_type

    if invalid_send_type:
        error_sheets["Invalid_Send_Types"] = pd.DataFrame(
            {
                "Message Name": invalid_send_type.keys(),
                "Invalid Send Type": invalid_send_type.values(),
            }
        )

    # 5. Frame Format errors (only for CANFD)
    if "Frame Format" in data_frame.columns:
        msg_frame_format = dict(zip(data_frame["Msg Name"], data_frame["Frame Format"]))
        invalid_ff = {}
        for mes, ff in msg_frame_format.items():
            if ff not in ["StandardCAN_FD", "StandardCAN"]:
                invalid_ff[mes] = ff

        if invalid_ff:
            error_sheets["Invalid_Frame_Formats"] = pd.DataFrame(
                {
                    "Message Name": invalid_ff.keys(),
                    "Invalid Frame Format": invalid_ff.values(),
                }
            )

    # 6. BRS errors (only for CANFD)
    if "BRS" in data_frame.columns:
        msg_brs = dict(zip(data_frame["Msg Name"], data_frame["BRS"]))
        msg_frame_format = dict(zip(data_frame["Msg Name"], data_frame["Frame Format"]))
        invalid_brs = {}
        invalid_brs_protocol = {}
        for mes, brs in msg_brs.items():
            if brs not in [0, 1]:
                invalid_brs[mes] = brs
            if brs == 0 and msg_frame_format[mes] != "StandardCAN":
                invalid_brs_protocol[mes] = {
                    "BRS": brs,
                    "Frame Format": msg_frame_format[mes],
                }
            if brs == 1 and msg_frame_format[mes] != "StandardCAN_FD":
                invalid_brs_protocol[mes] = {
                    "BRS": brs,
                    "Frame Format": msg_frame_format[mes],
                }

        if invalid_brs:
            error_sheets["Invalid_BRS_Values"] = pd.DataFrame(
                {
                    "Message Name": invalid_brs.keys(),
                    "Invalid BRS": invalid_brs.values(),
                }
            )
        if invalid_brs_protocol:
            df_data = []
            for msg_name, values in invalid_brs_protocol.items():
                df_data.append(
                    {
                        "Message Name": msg_name,
                        "BRS": values["BRS"],
                        "Frame Format": values["Frame Format"],
                    }
                )
            error_sheets["BRS_Frame_Format_Mismatch"] = pd.DataFrame(df_data)

    # 7. Message Length errors
    if "Frame Format" in data_frame.columns:
        msg_len = dict(zip(data_frame["Msg Name"], data_frame["Length"]))
        msg_frame_format = dict(zip(data_frame["Msg Name"], data_frame["Frame Format"]))
        invalid_len = {}
        for mes, length in msg_len.items():
            if length == 8 and msg_frame_format[mes] != "StandardCAN":
                invalid_len[mes] = {
                    "Length": length,
                    "Frame Format": msg_frame_format[mes],
                }
            if (length == 64 or length == 8) and msg_frame_format[
                mes
            ] != "StandardCAN_FD":
                invalid_len[mes] = {
                    "Length": length,
                    "Frame Format": msg_frame_format[mes],
                }

        if invalid_len:
            df_data = []
            for msg_name, values in invalid_len.items():
                df_data.append(
                    {
                        "Message Name": msg_name,
                        "Length": values["Length"],
                        "Frame Format": values["Frame Format"],
                    }
                )
            error_sheets["Invalid_Message_Lengths"] = pd.DataFrame(df_data)

    # 8. Signal Name errors
    invalid_sig_names = []
    too_long_sig_names = []
    need_change_sig_names = []
    sig_name = set(data_frame["Sig Name"].dropna().astype(str))
    for name in sig_name:
        if not re.fullmatch(r"^[A-Za-z0-9_\-]+$", name.strip()):
            invalid_sig_names.append(name)
        if len(name) > 64:
            too_long_sig_names.append(name)
        if len(name) > 36 and len(name) < 64:
            need_change_sig_names.append(name)

    if invalid_sig_names:
        error_sheets["Invalid_Signal_Names"] = pd.DataFrame(
            {"Incorrect name": invalid_sig_names}
        )
    if too_long_sig_names:
        error_sheets["Long_Signal_Names"] = pd.DataFrame(
            {"Name": too_long_sig_names, "Length": [len(n) for n in too_long_sig_names]}
        )
    if need_change_sig_names:
        error_sheets["Signal_Names_Need_Shortening"] = pd.DataFrame(
            {
                "Name": need_change_sig_names,
                "Length": [len(n) for n in need_change_sig_names],
            }
        )

    # 9. Signal Value Description errors
    sig_desc = dict(zip(data_frame["Sig Name"], data_frame["Signal Value Description"]))
    invalid_val = {}
    invalid_nan = {}
    for mes, val in sig_desc.items():
        if pd.isna(val):
            invalid_nan[mes] = "NaN"
            continue
        str_val = str(val)
        if not re.fullmatch(r"^[A-Za-z0-9 ,.;]+$", str_val):
            invalid_val[mes] = str_val

    if invalid_val:
        error_sheets["Invalid_Signal_Value_Descriptions"] = pd.DataFrame(
            {
                "Signal Name": invalid_val.keys(),
                "Invalid Description": invalid_val.values(),
            }
        )
    if invalid_nan:
        error_sheets["Missing_Signal_Value_Descriptions"] = pd.DataFrame(
            {"Signal Name": invalid_nan.keys(), "Status": invalid_nan.values()}
        )

    # 10. Signal Description errors
    sig_desc = dict(zip(data_frame["Sig Name"], data_frame["Description"]))
    invalid_val = {}
    invalid_nan = {}
    for mes, val in sig_desc.items():
        if pd.isna(val):
            invalid_nan[mes] = "NaN"
            continue
        str_val = str(val)
        if not re.fullmatch(r"^[A-Za-z0-9 ,.;:+_/-<>%()°~-]+$", str_val):
            invalid_val[mes] = str_val

    if invalid_val:
        error_sheets["Invalid_Signal_Descriptions"] = pd.DataFrame(
            {
                "Signal Name": invalid_val.keys(),
                "Invalid Description": invalid_val.values(),
            }
        )
    if invalid_nan:
        error_sheets["Missing_Signal_Descriptions"] = pd.DataFrame(
            {"Signal Name": invalid_nan.keys(), "Status": invalid_nan.values()}
        )

    # 11. Byte Order errors
    byte_order = dict(zip(data_frame["Sig Name"], data_frame["Byte Order"]))
    invalid_order = {}
    for mes, byte in byte_order.items():
        if byte != "Motorola MSB":
            invalid_order[mes] = byte

    if invalid_order:
        error_sheets["Invalid_Byte_Orders"] = pd.DataFrame(
            {"Signal Name": invalid_order.keys(), "Byte Order": invalid_order.values()}
        )

    # 12. Start Byte errors
    start_byte = dict(zip(data_frame["Sig Name"], data_frame["Start Byte"]))
    invalid_byte = {}
    for sig, byte in start_byte.items():
        if byte not in range(0, 8):
            invalid_byte[sig] = byte

    if invalid_byte:
        error_sheets["Invalid_Start_Bytes"] = pd.DataFrame(
            {"Signal Name": invalid_byte.keys(), "Start Byte": invalid_byte.values()}
        )

    # 13. Start Bit errors
    start_bit = dict(zip(data_frame["Sig Name"], data_frame["Start Bit"]))
    invalid_bit = {}
    for sig, byte in start_bit.items():
        if byte not in range(0, 64):
            invalid_bit[sig] = byte

    if invalid_bit:
        error_sheets["Invalid_Start_Bits"] = pd.DataFrame(
            {"Signal Name": invalid_bit.keys(), "Start Bit": invalid_bit.values()}
        )

    # 14. Signal Send Type errors
    sig_send_type = dict(zip(data_frame["Sig Name"], data_frame["Signal Send Type"]))
    msg_send_type = dict(zip(data_frame["Msg Name"], data_frame["Send Type"]))
    validation_rules = {
        "CA": ["Cycle", "IfActiveWithRepetition"],
        "CE": [
            "Cycle",
            "OnWrite",
            "OnChange",
            "OnWriteWithRepetition",
            "OnChangeWithRepetition",
        ],
        "Cycle": ["Cycle"],
        "Event": [
            "OnWrite",
            "OnChange",
            "OnWriteWithRepetition",
            "OnChangeWithRepetition",
        ],
        "IfActive": ["IfActive"],
    }
    invalid_signals = []
    for sig_name, sig_type in sig_send_type.items():
        msg_name = data_frame[data_frame["Sig Name"] == sig_name]["Msg Name"].iloc[0]
        msg_type = msg_send_type.get(msg_name)
        if msg_type in validation_rules:
            allowed_types = validation_rules[msg_type]
            if sig_type not in allowed_types:
                invalid_signals.append(
                    {
                        "Signal Name": sig_name,
                        "Message Name": msg_name,
                        "Message Send Type": msg_type,
                        "Signal Send Type": sig_type,
                        "Expected Types": ", ".join(allowed_types),
                    }
                )

    if invalid_signals:
        error_sheets["Invalid_Signal_Send_Types"] = pd.DataFrame(invalid_signals)

    # 15. Resolution errors
    resol = dict(zip(data_frame["Sig Name"], data_frame["Resolution"]))
    invalid_type = {}
    invalid_val = {}
    for sig, res in resol.items():
        if pd.isna(res):
            invalid_val[sig] = res
            continue
        if type(res) not in [int, float]:
            invalid_type[sig] = res

    if invalid_type:
        error_sheets["Invalid_Resolution_Types"] = pd.DataFrame(
            {"Signal Name": invalid_type.keys(), "Resolution": invalid_type.values()}
        )
    if invalid_val:
        error_sheets["Missing_Resolutions"] = pd.DataFrame(
            {"Signal Name": invalid_val.keys(), "Status": invalid_val.values()}
        )

    # 16. Offset errors
    offset = dict(zip(data_frame["Sig Name"], data_frame["Offset"]))
    invalid_val = {}
    invalid_type = {}
    for sig, off in offset.items():
        if pd.isna(off):
            invalid_val[sig] = off
            continue
        if type(off) not in [int, float]:
            invalid_type[sig] = off

    if invalid_type:
        error_sheets["Invalid_Offset_Types"] = pd.DataFrame(
            {"Signal Name": invalid_type.keys(), "Offset": invalid_type.values()}
        )
    if invalid_val:
        error_sheets["Missing_Offsets"] = pd.DataFrame(
            {"Signal Name": invalid_val.keys(), "Status": invalid_val.values()}
        )

    # 17. Minimum value errors
    min_phys = dict(zip(data_frame["Sig Name"], data_frame["Min"]))
    min_hex = dict(zip(data_frame["Sig Name"], data_frame["Min Hex"]))
    resolutions = dict(zip(data_frame["Sig Name"], data_frame["Resolution"]))
    offsets = dict(zip(data_frame["Sig Name"], data_frame["Offset"]))
    invalid_signals = []
    for sig_name in min_phys.keys():
        phys = min_phys.get(sig_name)
        hex_val = min_hex.get(sig_name)
        res = resolutions.get(sig_name)
        offset = offsets.get(sig_name, 0)

        if pd.isna(phys) or pd.isna(hex_val) or pd.isna(res):
            continue

        try:
            if isinstance(hex_val, str):
                if hex_val.startswith("0x"):
                    hex_int = int(hex_val, 16)
                else:
                    hex_int = int(hex_val)
            else:
                hex_int = int(hex_val)

            calculated_phys = hex_int * res + offset

            if not math.isclose(calculated_phys, phys, rel_tol=1e-9):
                invalid_signals.append(
                    {
                        "Signal Name": sig_name,
                        "Min (Physical)": phys,
                        "Min (Hex)": hex_val,
                        "Calculated Physical": calculated_phys,
                        "Resolution": res,
                        "Offset": offset,
                        "Difference": abs(calculated_phys - phys),
                    }
                )

        except (ValueError, TypeError) as e:
            invalid_signals.append(
                {
                    "Signal Name": sig_name,
                    "Error": f"Invalid data format: {str(e)}",
                    "Min (Hex)": hex_val,
                    "Resolution": res,
                    "Offset": offset,
                }
            )

    if invalid_signals:
        error_sheets["Invalid_Minimum_Values"] = pd.DataFrame(invalid_signals)

    # 18. Maximum value errors
    max_phys = dict(zip(data_frame["Sig Name"], data_frame["Max"]))
    max_hex = (
        dict(zip(data_frame["Sig Name"], data_frame["Max Hex"]))
        if "Max Hex" in data_frame.columns
        else {}
    )
    if not max_hex:
        max_hex = dict(zip(data_frame["Sig Name"], data_frame["Invalid"]))
    resolutions = dict(zip(data_frame["Sig Name"], data_frame["Resolution"]))
    offsets = dict(zip(data_frame["Sig Name"], data_frame["Offset"]))
    invalid_signals = []
    for sig_name in max_phys.keys():
        phys = max_phys.get(sig_name)
        hex_val = max_hex.get(sig_name)
        res = resolutions.get(sig_name)
        offset = offsets.get(sig_name, 0)

        if pd.isna(phys) or pd.isna(hex_val) or pd.isna(res):
            continue

        try:
            if isinstance(hex_val, str):
                if hex_val.startswith("0x"):
                    hex_int = int(hex_val, 16)
                else:
                    hex_int = int(hex_val)
            else:
                hex_int = int(hex_val)

            calculated_phys = hex_int * res + offset

            if not math.isclose(calculated_phys, phys, rel_tol=1e-9):
                invalid_signals.append(
                    {
                        "Signal Name": sig_name,
                        "Max (Physical)": phys,
                        "Max (Hex)": hex_val,
                        "Calculated Physical": calculated_phys,
                        "Resolution": res,
                        "Offset": offset,
                        "Difference": abs(calculated_phys - phys),
                    }
                )

        except (ValueError, TypeError) as e:
            invalid_signals.append(
                {
                    "Signal Name": sig_name,
                    "Error": f"Invalid data format: {str(e)}",
                    "Max (Hex)": hex_val,
                    "Resolution": res,
                    "Offset": offset,
                }
            )

    if invalid_signals:
        error_sheets["Invalid_Maximum_Values"] = pd.DataFrame(invalid_signals)

    # Create the Excel file
    if error_sheets:
        with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
            for sheet_name, df in error_sheets.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        return True
    else:
        return False


def validate_messages_name(data_frame: pd.DataFrame) -> bool:
    invalid_names = []
    too_long_names = []

    msg_names = set(data_frame["Msg Name"].dropna().astype(str))

    for name in msg_names:
        if not re.fullmatch(r"^[A-Za-z0-9_\-]+$", name.strip()):
            invalid_names.append(name)

        if len(name) > 64:
            too_long_names.append(name)

    if not invalid_names and not too_long_names:
        st.success("All message titles are correct!")
        return True

    if invalid_names:
        with st.expander(
            "Incorrect names (contain prohibited characters)", expanded=True
        ):
            st.error(f"Found {len(invalid_names)} incorrect name:")
            st.dataframe(pd.DataFrame({"Incorrect name": invalid_names}))
            st.info("Allowed characters: A-Z, a-z, 0-9, _, -")

    if too_long_names:
        with st.expander("Names too long (>64 characters)", expanded=True):
            st.warning(f"Found {len(too_long_names)} too long name:")
            st.dataframe(
                pd.DataFrame(
                    {"Name": too_long_names, "Len": [len(n) for n in too_long_names]}
                )
            )

    return False


def validate_messages_type(data_frame: pd.DataFrame) -> bool:
    msg_type = dict(zip(data_frame["Msg Name"], data_frame["Msg Type"]))

    invalid_type = {}
    invalid_name = {}

    for key, val in msg_type.items():
        if val not in ["Normal", "Diag", "NM"]:
            invalid_type[key] = val

        if key.startswith("Diag") and val != "Diag":
            invalid_name[key] = val

        if key.startswith("NM_") and val != "NM":
            invalid_name[key] = val

    if not invalid_name and not invalid_type:
        st.success("All message types are correct!")
        return True

    if invalid_type:
        with st.expander("Incorrect type (Unknown type)", expanded=True):
            st.error(f"Found {len(invalid_type.keys())} incorrect type:")
            st.dataframe(
                pd.DataFrame(
                    {
                        "Mes Name": invalid_type.keys(),
                        "Incorrect types": invalid_type.values(),
                    }
                )
            )
            st.info("list of allowed values ​​'Normal', 'Diag', 'NM'")

    if invalid_name:
        with st.expander("Incorrect name (Not for this type))", expanded=True):
            st.error(f"Found {len(invalid_name.keys())} incorrect type:")
            st.dataframe(
                pd.DataFrame(
                    {
                        "Incorrect Name": invalid_name.keys(),
                        "Msg Type": invalid_name.values(),
                    }
                )
            )
            st.info(
                "NM, if Msg Name first 3 characters = 'NM_' and Diag, if Msg Name firsts 4 characters = 'Diag'"
            )

    return False


def validate_messages_id(data_frame: pd.DataFrame) -> bool:
    data_frame["Msg ID"] = data_frame["Msg ID"].apply(
        lambda x: int(x, 16) if isinstance(x, str) and x.startswith("0x") else int(x)
    )

    msg_id = dict(zip(data_frame["Msg Name"], data_frame["Msg ID"]))
    msg_type = dict(zip(data_frame["Msg Name"], data_frame["Msg Type"]))

    invalid_id = {}
    invalid_type = {}

    for mes, id in msg_id.items():
        if not (0x001 <= id <= 0x7FF):
            invalid_id[mes] = id
        if 0x700 <= id <= 0x7FF and msg_type[mes] != "Diag":
            invalid_type[mes] = id
        if 0x500 <= id <= 0x5FF and msg_type[mes] != "NM":
            invalid_type[mes] = id

    if not invalid_type and not invalid_id:
        st.success("All message IDs are correct!")
        return True

    if invalid_id:
        with st.expander(
            "Incorrect ID (Whether it fits within the range or not))", expanded=True
        ):
            st.error(f"Found {len(invalid_id.keys())} incorrect IDs:")
            st.dataframe(
                pd.DataFrame(
                    {
                        "Msg Name": invalid_id.keys(),
                        "Incorrect IDs": invalid_id.values(),
                    }
                )
            )
            st.info("Msg ID - Must be in the range 0x001 to 0x7FF (Hex)")

    if invalid_type:
        with st.expander("Incorrect ID for Msg Type (Wrong range)", expanded=True):
            st.error(f"Found {len(invalid_type.keys())} incorrect types:")
            st.dataframe(
                pd.DataFrame(
                    {
                        "Msg Name": invalid_type.keys(),
                        "Incorrect IDs": invalid_type.values(),
                    }
                )
            )
            st.info(
                "Diag if Message ID is in the range 0x700 to 7FF and NM if Message ID is in the range 0x500 to 5FF"
            )

    return False


def validate_messages_send_type(data_frame: pd.DataFrame) -> bool:

    msg_send_type = dict(zip(data_frame["Msg Name"], data_frame["Send Type"]))

    invalid_send_type = {}

    for mes, send_type in msg_send_type.items():
        if send_type not in ["Cycle", "Event", "CE"]:
            invalid_send_type[mes] = send_type

    if not invalid_send_type:
        st.success("All messages send types are correct!")
        return True

    if invalid_send_type:
        with st.expander("Incorrect send types()", expanded=True):
            st.error(f"Found {len(invalid_send_type.keys())} incorrect send types:")
            st.dataframe(
                pd.DataFrame(
                    {
                        "Msg Name": invalid_send_type.keys(),
                        "Incorrect type": invalid_send_type.values(),
                    }
                )
            )
            st.info("Send Type should be 'Cycle', 'Event' or 'CE'")

    return False


def validate_messages_frame_fromat(
    file_path: Union[UploadedFile, str, List], data_frame: pd.DataFrame
) -> bool:
    file_info = get_file_info(file_path)

    if file_info["protocol"] != "CANFD":
        st.warning(
            f"Frame Format validation is not applicable for {file_info['protocol']} protocol"
        )
        return True

    if "Frame Format" not in data_frame.columns:
        st.error("Frame Format column not found in the dataframe")
        return False
    msg_frame_format = dict(zip(data_frame["Msg Name"], data_frame["Frame Format"]))

    invalid_ff = {}

    for mes, ff in msg_frame_format.items():
        if ff not in ["StandardCAN_FD", "StandardCAN"]:
            invalid_ff[mes] = ff

    if not invalid_ff:
        st.success("All messages frame formats are correct!")
        return True

    if invalid_ff:
        st.error(f"Found {len(invalid_ff.keys())} incorrect frame formats:")
        st.dataframe(
            pd.DataFrame(
                {
                    "Msg Name": invalid_ff.keys(),
                    "Incorrect frame format": invalid_ff.values(),
                }
            )
        )
        st.info("Frame format should be 'StandardCAN_FD' or 'StandardCAN'")

    return False


def validate_messages_BRS(
    file_path: Union[UploadedFile, str, List], data_frame: pd.DataFrame
) -> bool:
    file_info = get_file_info(file_path)

    if file_info["protocol"] != "CANFD":
        st.warning(
            f"BRS validation is not applicable for {file_info['protocol']} protocol"
        )
        return True

    if "BRS" not in data_frame.columns:
        st.error("BRS column not found in the dataframe")
        return False

    msg_brs = dict(zip(data_frame["Msg Name"], data_frame["BRS"]))
    msg_frame_format = dict(zip(data_frame["Msg Name"], data_frame["Frame Format"]))

    invalid_brs = {}
    invalid_brs_protocol = {}

    for mes, brs in msg_brs.items():
        if brs not in [0, 1]:
            invalid_brs[mes] = brs

        if brs == 0 and msg_frame_format[mes] != "StandardCAN":
            invalid_brs_protocol[mes] = {
                "BRS": brs,
                "Frame Format": msg_frame_format[mes],
            }

        if brs == 1 and msg_frame_format[mes] != "StandardCAN_FD":
            invalid_brs_protocol[mes] = {
                "BRS": brs,
                "Frame Format": msg_frame_format[mes],
            }

    if not invalid_brs and not invalid_brs_protocol:
        st.success("All BRS values are correct!")
        return True

    if invalid_brs:
        with st.expander("Incorrect BRS values", expanded=True):
            st.error(f"Found {len(invalid_brs)} incorrect BRS values:")
            st.dataframe(
                pd.DataFrame(
                    {
                        "Msg Name": invalid_brs.keys(),
                        "Incorrect BRS": invalid_brs.values(),
                    }
                )
            )
            st.info("BRS should be '1' or '0'")

    if invalid_brs_protocol:
        with st.expander("Incorrect BRS for Frame Format", expanded=True):
            st.error(
                f"Found {len(invalid_brs_protocol)} incorrect BRS for Frame Format"
            )
            df_data = []
            for msg_name, values in invalid_brs_protocol.items():
                df_data.append(
                    {
                        "Msg Name": msg_name,
                        "Incorrect BRS": values["BRS"],
                        "Frame Format": values["Frame Format"],
                    }
                )
            st.dataframe(pd.DataFrame(df_data))
            st.info(
                "BRS=0 should be with StandardCAN, BRS=1 should be with StandardCAN_FD"
            )

    return False


def validate_messages_length(data_frame: pd.DataFrame) -> bool:
    msg_len = dict(zip(data_frame["Msg Name"], data_frame["Length"]))
    msg_frame_format = dict(zip(data_frame["Msg Name"], data_frame["Frame Format"]))

    invalid_len = {}

    for mes, length in msg_len.items():
        frame_format = msg_frame_format[mes] 

        if frame_format == "StandardCAN" and length != 8:
            invalid_len[mes] = {"Len": length, "Frame": frame_format}

        elif frame_format == "StandardCAN_FD" and length not in (8, 64):
            invalid_len[mes] = {"Len": length, "Frame": frame_format}

        elif frame_format not in ("StandardCAN", "StandardCAN_FD"):
            invalid_len[mes] = {"Len": length, "Frame": frame_format}

    if not invalid_len:
        st.success("All messages length are correct!")
        return True

    if invalid_len:
        with st.expander("Incorrect messages length for frame format", expanded=True):
            st.error(f"Found {len(invalid_len.keys())} incorrect length")
            df_data = []
            for msg_name, values in invalid_len.items():
                df_data.append(
                    {
                        "Msg Name": msg_name,
                        "Incorrect Length": values["Len"],
                        "Frame Format": values["Frame"],
                    }
                )
            st.dataframe(pd.DataFrame(df_data))
            st.info(
                "For StandardCAN, message length must be 8 bytes. "
                "For StandardCAN_FD, message length must be 8 or 64 bytes."
            )

    return False

def validate_signal_names(data_frame: pd.DataFrame) -> bool:
    invalid_names = []
    too_long_names = []
    need_change = []
    sig_name = set(data_frame["Sig Name"].dropna().astype(str))

    for name in sig_name:
        if not re.fullmatch(r"^[A-Za-z0-9_\-]+$", name.strip()):
            invalid_names.append(name)

        if len(name) > 64:
            too_long_names.append(name)
        if len(name) > 36 and len(name) < 64:
            need_change.append(name)

    if not invalid_names and not too_long_names and not need_change:
        st.success("All signals titles are correct!")
        return True

    if invalid_names:
        with st.expander(
            "Incorrect names (contain prohibited characters)", expanded=True
        ):
            st.error(f"Found {len(invalid_names)} incorrect name:")
            st.dataframe(pd.DataFrame({"Incorrect name": invalid_names}))
            st.info("Allowed characters: A-Z, a-z, 0-9, _, -")

    if too_long_names:
        with st.expander("Names too long (>64 characters)", expanded=True):
            st.warning(f"Found {len(too_long_names)} too long name:")
            st.dataframe(
                pd.DataFrame(
                    {"Name": too_long_names, "Len": [len(n) for n in too_long_names]}
                )
            )

    if need_change:
        with st.expander("Names, which need change (>36 characters)", expanded=True):
            st.warning(f"Found {len(need_change)} need change name:")
            st.dataframe(
                pd.DataFrame(
                    {"Name": need_change, "Len": [len(n) for n in need_change]}
                )
            )
            st.info("Please, try to make the Signal name shorter")

    return False


def validate_signal_value_description(data_frame: pd.DataFrame) -> bool:
    sig_desc = dict(zip(data_frame["Sig Name"], data_frame["Signal Value Description"]))

    invalid_val = {}
    invalid_nan = {}

    for mes, val in sig_desc.items():
        if pd.isna(val):
            invalid_nan[mes] = "NaN"
            continue

        str_val = str(val)
        if not re.fullmatch(r"^[A-Za-z0-9 ,.;]+$", str_val):
            invalid_val[mes] = str_val

    if not invalid_nan and not invalid_val:
        st.success("All Signal Values Description are correct!")
        return True

    if invalid_val:
        with st.expander(
            "Incorrect signal value description (not needed characters)", expanded=True
        ):
            st.error(
                f"Found {len(invalid_val)} value descriptions with wrong characters"
            )
            st.dataframe(
                pd.DataFrame(
                    {
                        "Signal Name": invalid_val.keys(),
                        "Incorrect Description": invalid_val.values(),
                    }
                )
            )
            st.info(
                "Allowed characters: A-Z, a-z, 0-9, spaces, commas, periods, and semicolons"
            )

    if invalid_nan:
        with st.expander("NaN Value Description", expanded=True):
            st.error(f"Found {len(invalid_nan)} signals without value description")
            st.dataframe(
                pd.DataFrame(
                    {
                        "Signal Name": invalid_nan.keys(),
                        "Description Status": invalid_nan.values(),
                    }
                )
            )

    return False if (invalid_nan or invalid_val) else True


def validate_signal_descriprion(data_frame: pd.DataFrame) -> bool:
    sig_desc = dict(zip(data_frame["Sig Name"], data_frame["Description"]))

    invalid_val = {}
    invalid_nan = {}

    for mes, val in sig_desc.items():
        if pd.isna(val):
            invalid_nan[mes] = "NaN"
            continue

        str_val = str(val)
        if not re.fullmatch(r"^[A-Za-z0-9 ,.;:+_/-<>%()°~-]+$", str_val):
            invalid_val[mes] = str_val

    if not invalid_nan and not invalid_val:
        st.success("All Signal Description are correct!")
        return True

    if invalid_val:
        with st.expander(
            "Incorrect signal description (not needed characters)", expanded=True
        ):
            st.error(
                f"Found {len(invalid_val)} value descriptions with wrong characters"
            )
            st.dataframe(
                pd.DataFrame(
                    {
                        "Signal Name": invalid_val.keys(),
                        "Incorrect Description": invalid_val.values(),
                    }
                )
            )
            st.info(
                "Allowed characters: A-Z, a-z, 0-9, spaces, commas, periods, and semicolons"
            )

    if invalid_nan:
        with st.expander("NaN Value Description", expanded=True):
            st.error(f"Found {len(invalid_nan)} signals without value description")
            st.dataframe(
                pd.DataFrame(
                    {
                        "Signal Name": invalid_nan.keys(),
                        "Description Status": invalid_nan.values(),
                    }
                )
            )

    return False if (invalid_nan or invalid_val) else True


def validate_byte_order(data_frame: pd.DataFrame) -> bool:

    byte_order = dict(zip(data_frame["Sig Name"], data_frame["Byte Order"]))

    invalid_order = {}

    for mes, byte in byte_order.items():
        if byte != "Motorola MSB":
            invalid_order[mes] = byte

    if not invalid_order:
        st.success("All Signal Byte Orders are correct!")

    if invalid_order:
        with st.expander("Incorrect Byte Order", expanded=True):
            st.error(f"Found {len(invalid_order.keys())} incorrect byte order")
            st.dataframe(
                pd.DataFrame(
                    {
                        "Signal Name": invalid_order.keys(),
                        "Incorrect Byte Order": invalid_order.values(),
                    }
                )
            )
            st.info("Byte Order in valid value 'Motorola MSB'")

    return False


def validate_start_byte(data_frame: pd.DataFrame) -> bool:
    start_byte = dict(zip(data_frame["Sig Name"], data_frame["Start Byte"]))

    invalid_byte = {}

    for sig, byte in start_byte.items():
        if byte not in range(0, 8):
            invalid_byte[sig] = byte

    if not invalid_byte:
        st.success("All Start Byte are correct!")
        return True

    if invalid_byte:
        with st.expander("Inccorect Start Byte", expanded=True):
            st.error(f"Found {len(invalid_byte.keys())} incorrect start byte")
            st.dataframe(
                pd.DataFrame(
                    {
                        "Signal Name": invalid_byte.keys(),
                        "Incorrect Start Byte": invalid_byte.values(),
                    }
                )
            )
            st.info("Start Byte is only a number, in the range from 0 to 7")

    return False


def validate_start_bit(data_frame: pd.DataFrame) -> bool:
    start_bit = dict(zip(data_frame["Sig Name"], data_frame["Start Bit"]))

    invalid_bit = {}

    for sig, byte in start_bit.items():
        if byte not in range(0, 64):
            invalid_bit[sig] = byte

    if not invalid_bit:
        st.success("All Start Bit are correct!")
        return True

    if invalid_bit:
        with st.expander("Inccorect Start Bit", expanded=True):
            st.error(f"Found {len(invalid_bit.keys())} incorrect start bit")
            st.dataframe(
                pd.DataFrame(
                    {
                        "Signal Name": invalid_bit.keys(),
                        "Incorrect Start Bit": invalid_bit.values(),
                    }
                )
            )
            st.info("Start Bit is only a number, in the range from 0 to 63")

    return False


def validate_signal_send_type(data_frame: pd.DataFrame) -> bool:
    sig_send_type = dict(zip(data_frame["Sig Name"], data_frame["Signal Send Type"]))
    msg_send_type = dict(zip(data_frame["Msg Name"], data_frame["Send Type"]))

    validation_rules = {
        "CA": ["Cycle", "IfActiveWithRepetition"],
        "CE": [
            "Cycle",
            "OnWrite",
            "OnChange",
            "OnWriteWithRepetition",
            "OnChangeWithRepetition",
        ],
        "Cycle": ["Cycle"],
        "Event": [
            "OnWrite",
            "OnChange",
            "OnWriteWithRepetition",
            "OnChangeWithRepetition",
        ],
        "IfActive": ["IfActive"],
    }

    invalid_signals = []

    for sig_name, sig_type in sig_send_type.items():
        msg_name = data_frame[data_frame["Sig Name"] == sig_name]["Msg Name"].iloc[0]
        msg_type = msg_send_type.get(msg_name)

        if msg_type in validation_rules:
            allowed_types = validation_rules[msg_type]
            if sig_type not in allowed_types:
                invalid_signals.append(
                    {
                        "Signal Name": sig_name,
                        "Message Name": msg_name,
                        "Message Send Type": msg_type,
                        "Signal Send Type": sig_type,
                        "Expected Types": ", ".join(allowed_types),
                    }
                )

    if not invalid_signals:
        st.success("All Signal Send Types are correct!")
        return True
    else:
        with st.expander("Invalid Signal Send Types", expanded=True):
            st.error(f"Found {len(invalid_signals)} invalid signal send types:")
            st.dataframe(pd.DataFrame(invalid_signals))
            st.info(
                """
            Validation rules:
            - If Msg Send Type == 'CA': Signal Send Type must be in ['Cycle', 'IfActiveWithRepetition']
            - If Msg Send Type == 'CE': Signal Send Type must be in ['Cycle', 'OnWrite', 'OnChange', 'OnWriteWithRepetition', 'OnChangeWithRepetition']
            - If Msg Send Type == 'Cycle': Signal Send Type must be 'Cycle'
            - If Msg Send Type == 'Event': Signal Send Type must be in ['OnWrite', 'OnChange', 'OnWriteWithRepetition', 'OnChangeWithRepetition']
            - If Msg Send Type == 'IfActive': Signal Send Type must be 'IfActive'
            """
            )
        return False


def validate_resolution(data_frame: pd.DataFrame) -> bool:
    resol = dict(zip(data_frame["Sig Name"], data_frame["Resolution"]))

    invalid_type = {}
    invalid_val = {}

    for sig, res in resol.items():
        if pd.isna(res):
            invalid_val[sig] = res
            continue

        if type(res) not in [int, float]:
            invalid_type[sig] = res

    if not invalid_type and not invalid_val:
        st.success("All Resolutions are correct!")
        return True

    if invalid_type:
        with st.expander("Incorrect Resolution type", expanded=True):
            st.error(f"Found {len(invalid_type.keys())} inccorect resolution type")
            st.dataframe(
                pd.DataFrame(
                    {
                        "Sig Name": invalid_type.keys(),
                        "Incorrect Resolution": invalid_type.values(),
                    }
                )
            )
            st.info("Resolution is only int or float")

    if invalid_val:
        with st.expander("Incorrect Resolution value", expanded=True):
            st.error(f"Found {len(invalid_val.keys())} incorrect resolution value")
            st.dataframe(
                pd.DataFrame(
                    {
                        "Sig Name": invalid_val.keys(),
                        "Incorrect Resolution": invalid_val.values(),
                    }
                )
            )
            st.info("Resolution must be in matrix")

    return False


def validate_offset(data_frame: pd.DataFrame) -> bool:
    offset = dict(zip(data_frame["Sig Name"], data_frame["Offset"]))

    invalid_val = {}
    invalid_type = {}

    for sig, off in offset.items():
        if pd.isna(off):
            invalid_val[sig] = off
            continue

        if type(off) not in [int, float]:
            invalid_type[sig] = off

    if not invalid_val and not invalid_type:
        st.success("All Signals Offset are correct!")
        return True

    if invalid_type:
        with st.expander("Incorrect Offset type", expanded=True):
            st.error(f"Found {len(invalid_type.keys())} incorrect offset type")
            st.dataframe(
                pd.DataFrame(
                    {
                        "Sig Name": invalid_type.keys(),
                        "Incorrect Offset": invalid_type.values(),
                    }
                )
            )
            st.info("Offset must be int or float")

    if invalid_val:
        with st.expander("Incorrect Offset value", expanded=True):
            st.error(f"Found {len(invalid_val.keys())} incorrect offset value")
            st.dataframe(
                pd.DataFrame(
                    {
                        "Sig Name": invalid_val.keys(),
                        "Incorrect Offset": invalid_val.values(),
                    }
                )
            )
            st.info("Offset must be in matrix")

    return False


def validate_minimum(data_frame: pd.DataFrame) -> bool:
    min_phys = dict(zip(data_frame["Sig Name"], data_frame["Min"]))
    min_hex = dict(zip(data_frame["Sig Name"], data_frame["Min Hex"]))
    resolutions = dict(zip(data_frame["Sig Name"], data_frame["Resolution"]))
    offsets = dict(zip(data_frame["Sig Name"], data_frame["Offset"]))

    invalid_signals = []

    for sig_name in min_phys.keys():
        phys = min_phys.get(sig_name)
        hex_val = min_hex.get(sig_name)
        res = resolutions.get(sig_name)
        offset = offsets.get(sig_name, 0)

        if pd.isna(phys) or pd.isna(hex_val) or pd.isna(res):
            continue

        try:
            if isinstance(hex_val, str):
                if hex_val.startswith("0x"):
                    hex_int = int(hex_val, 16)
                else:
                    hex_int = int(hex_val)
            else:
                hex_int = int(hex_val)

            calculated_phys = hex_int * res + offset

            if not math.isclose(calculated_phys, phys, rel_tol=1e-9):
                invalid_signals.append(
                    {
                        "Signal Name": sig_name,
                        "Min (Physical)": phys,
                        "Min (Hex)": hex_val,
                        "Calculated Physical": calculated_phys,
                        "Resolution": res,
                        "Offset": offset,
                        "Difference": abs(calculated_phys - phys),
                    }
                )

        except (ValueError, TypeError) as e:
            invalid_signals.append(
                {
                    "Signal Name": sig_name,
                    "Error": f"Invalid data format: {str(e)}",
                    "Min (Hex)": hex_val,
                    "Resolution": res,
                    "Offset": offset,
                }
            )

    if not invalid_signals:
        st.success(
            "All minimum values match the formula: Physical = (Hex * Resolution) + Offset"
        )
        return True
    else:
        with st.expander("Invalid Minimum Values", expanded=True):
            st.error(
                f"Found {len(invalid_signals)} signals with incorrect minimum values"
            )

            df_errors = pd.DataFrame(invalid_signals)

            if "Error" in df_errors.columns:
                st.warning("Some values have format issues:")
                st.dataframe(
                    df_errors[df_errors["Error"].notna()][["Signal Name", "Error"]]
                )

                df_errors = df_errors[df_errors["Error"].isna()]

            if not df_errors.empty:
                st.dataframe(df_errors)
                st.info("Physical value should equal (Hex * Resolution) + Offset")

        return False


def validate_maximum(data_frame: pd.DataFrame) -> bool:
    max_phys = dict(zip(data_frame["Sig Name"], data_frame["Max"]))
    max_hex = (
        dict(zip(data_frame["Sig Name"], data_frame["Max Hex"]))
        if "Max Hex" in data_frame.columns
        else {}
    )
    resolutions = dict(zip(data_frame["Sig Name"], data_frame["Resolution"]))
    offsets = dict(zip(data_frame["Sig Name"], data_frame["Offset"]))

    if not max_hex:
        max_hex = dict(zip(data_frame["Sig Name"], data_frame["Invalid"]))

    invalid_signals = []

    for sig_name in max_phys.keys():
        phys = max_phys.get(sig_name)
        hex_val = max_hex.get(sig_name)
        res = resolutions.get(sig_name)
        offset = offsets.get(sig_name, 0)

        if pd.isna(phys) or pd.isna(hex_val) or pd.isna(res):
            continue

        try:
            if isinstance(hex_val, str):
                if hex_val.startswith("0x"):
                    hex_int = int(hex_val, 16)
                else:
                    hex_int = int(hex_val)
            else:
                hex_int = int(hex_val)

            calculated_phys = hex_int * res + offset

            if not math.isclose(calculated_phys, phys, rel_tol=1e-9):
                invalid_signals.append(
                    {
                        "Signal Name": sig_name,
                        "Max (Physical)": phys,
                        "Max (Hex)": hex_val,
                        "Calculated Physical": calculated_phys,
                        "Resolution": res,
                        "Offset": offset,
                        "Difference": abs(calculated_phys - phys),
                    }
                )

        except (ValueError, TypeError) as e:
            invalid_signals.append(
                {
                    "Signal Name": sig_name,
                    "Error": f"Invalid data format: {str(e)}",
                    "Max (Hex)": hex_val,
                    "Resolution": res,
                    "Offset": offset,
                }
            )

    if not invalid_signals:
        st.success(
            "All maximum values match the formula: Physical = (Hex * Resolution) + Offset"
        )
        return True
    else:
        with st.expander("Invalid Maximum Values", expanded=True):
            st.error(
                f"Found {len(invalid_signals)} signals with incorrect maximum values"
            )

            df_errors = pd.DataFrame(invalid_signals)

            if "Error" in df_errors.columns:
                st.warning("Some values have format issues:")
                st.dataframe(
                    df_errors[df_errors["Error"].notna()][["Signal Name", "Error"]]
                )

                df_errors = df_errors[df_errors["Error"].isna()]

            if not df_errors.empty:
                st.dataframe(df_errors)
                st.info("Physical value should equal (Hex * Resolution) + Offset")

        return False


def main():
    st.title("🚧CAN Messages Validator")
    uploaded_file = st.file_uploader("Upload matrix file", type=["xlsx"])

    if uploaded_file:
        try:
            df = load_xlsx(uploaded_file)
            processed_df = create_correct_df(df)

            st.success("File loaded successfully!")

            if st.button("Export All Validation Errors to Excel"):
                output_path = "validation_errors.xlsx"
                if export_validation_errors_to_excel(processed_df, output_path):
                    st.success(f"Validation errors exported to {output_path}")
                    with open(output_path, "rb") as f:
                        st.download_button(
                            label="Download Error Report",
                            data=f,
                            file_name=output_path,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        )
                else:
                    st.success("No validation errors found!")

            (
                tab1,
                tab2,
                tab3,
                tab4,
                tab5,
                tab6,
                tab7,
                tab8,
                tab9,
                tab10,
                tab11,
                tab12,
                tab13,
                tab14,
                tab15,
                tab16,
                tab17,
                tab18,
            ) = st.tabs(
                [
                    "Message Names",
                    "Message Types",
                    "Messages IDs",
                    "Messages Send Type",
                    "Messages Frame Format",
                    "Messages BRS",
                    "Messages Lenght",
                    "Signal Name",
                    "Signal Value Description",
                    "Signal Description",
                    "Byte Order",
                    "Start Byte",
                    "Start Bit",
                    "Signal Send Type",
                    "Resolution",
                    "Offset",
                    "Minimum",
                    "Maximum",
                ]
            )

            with tab1:
                validate_messages_name(processed_df)

            with tab2:
                validate_messages_type(processed_df)

            with tab3:
                validate_messages_id(processed_df)

            with tab4:
                validate_messages_send_type(processed_df)

            with tab5:
                validate_messages_frame_fromat(uploaded_file.name, processed_df)

            with tab6:
                validate_messages_BRS(uploaded_file.name, processed_df)

            with tab7:
                validate_messages_length(processed_df)

            with tab8:
                validate_signal_names(processed_df)

            with tab9:
                validate_signal_value_description(processed_df)

            with tab10:
                validate_signal_descriprion(processed_df)

            with tab11:
                validate_byte_order(processed_df)

            with tab12:
                validate_start_byte(processed_df)

            with tab13:
                validate_start_bit(processed_df)

            with tab14:
                validate_signal_send_type(processed_df)

            with tab15:
                validate_resolution(processed_df)

            with tab16:
                validate_offset(processed_df)

            with tab17:
                validate_minimum(processed_df)

            with tab18:
                validate_maximum(processed_df)

        except Exception as e:
            st.error(f"Error processing file: {str(e)}")
    else:
        st.info("Please upload an Excel file to begin validation")


if __name__ == "__main__":
    main()
