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


def get_engine(file_path: str) -> str:
    if isinstance(file_path, UploadedFile):
        if file_path.name.endswith(".xls"):
            return "xlrd"
        elif file_path.name.endswith((".xlsx", ".xlsm")):
            return "openpyxl"
        else:
            raise ValueError(f"Unsupported Excel file extension: {file_path.name}")
    else:
        if file_path.endswith(".xls"):
            return "xlrd"
        elif file_path.endswith((".xlsx", ".xlsm")):
            return "openpyxl"
        else:
            raise ValueError(f"Unsupported Excel file extension: {file_path}")


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
            engine = get_engine(file_path=file_path.name)
            data_frame = pd.read_excel(
                file_path, sheet_name="Matrix", keep_default_na=True, engine=engine
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
    # Identify bus users (nodes that send or receive messages)
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
        "Msg ID": df["Msg ID(hex)\n报文标识符"].ffill(),
        "Msg Name": df["Msg Name\n报文名称"].ffill(),
        "Protected ID": df["Protected ID (hex)\n保护标识符"].ffill(),
        "Send Type": df["Msg Send Type\n报文发送类型"].ffill(),
        "Checksum Mode": df["Checksum mode\n校验方式"].ffill(),
        "Msg Length": df["Msg Length(Byte)\n报文长度"].ffill(),
        "Sig Name": df["Signal Name\n信号名称"],
        "Description": df["Signal Description\n信号描述"],
        "Response Error": df["Response Error"],
        "Start Byte": df["Start Byte\n起始字节"],
        "Start Bit": df["Start Bit\n起始位"],
        "Length": df["Bit Length(Bit)\n信号长度"],
        "Resolution": df["Resolution\n精度"],
        "Offset": df["Offset\n偏移量"],
        "Min": df["Signal Min. Value(phys)\n物理最小值"],
        "Max": df["Signal Max. Value(phys)\n物理最大值"],
        "Min Hex": df["Signal Min. Value(Hex)\n总线最小值"],
        "Max Hex": df["Signal Max. Value(Hex)\n总线最大值"],
        "Unit": df["Unit\n单位"],
        "Initinal": df["Initial Value(Hex)\n初始值"],
        "Invalid": df["Invalid Value(Hex)\n无效值"],
        "Signal Value Description": df["Signal Value Description(hex)\n信号值描述"],
        "Remark": df["Remark\n备注"],
        "Receiver": receivers,
        "Senders": senders,
    }

    new_df = pd.DataFrame(new_df_data)

    new_df["Unit"] = new_df["Unit"].astype(str)
    new_df["Unit"] = new_df["Unit"].str.replace("Ω", "Ohm", regex=False)
    new_df["Unit"] = new_df["Unit"].str.replace("℃", "degC", regex=False)

    new_df = new_df.dropna(subset=["Sig Name"])

    new_df["Is Signed"] = False

    return new_df


def export_validation_errors_to_excel(data_frame: pd.DataFrame, file_path: str) -> bool:
    all_errors = []

    # 1. Message Name errors
    invalid_names = []
    too_long_names = []
    msg_names = set(data_frame["Msg Name"].dropna().astype(str))
    for name in msg_names:
        if not re.fullmatch(r"^[A-Za-z0-9_\-]+$", name.strip()):
            invalid_names.append(name)
        if len(name) > 32:
            too_long_names.append(name)

    for name in invalid_names:
        all_errors.append(
            {
                "Error Type": "Invalid Message Name",
                "Message/Signal Name": name,
                "Details": "Contains prohibited characters",
                "Expected": "Only A-Z, a-z, 0-9, _, - allowed",
            }
        )

    for name in too_long_names:
        all_errors.append(
            {
                "Error Type": "Too Long Message Name",
                "Message/Signal Name": name,
                "Details": f"Length: {len(name)} characters",
                "Expected": "Max 32 characters",
            }
        )

    # 2. Protected ID errors
    if "Protected ID" in data_frame.columns and "Msg ID" in data_frame.columns:
        try:
            data_frame["Protected ID"] = data_frame["Protected ID"].apply(
                lambda x: (
                    int(x, 16) if isinstance(x, str) and x.startswith("0x") else int(x)
                )
            )
            data_frame["Msg ID"] = data_frame["Msg ID"].apply(
                lambda x: (
                    int(x, 16) if isinstance(x, str) and x.startswith("0x") else int(x)
                )
            )

            prot_id = dict(zip(data_frame["Msg Name"], data_frame["Protected ID"]))
            msg_id = dict(zip(data_frame["Msg Name"], data_frame["Msg ID"]))

            invalid_range = {}
            invalid_calculation = {}
            invalid_parity = {}

            for mes, pid in prot_id.items():
                frame_id = msg_id.get(mes, -1)

                if not (0x00 <= pid <= 0xFF):
                    all_errors.append(
                        {
                            "Error Type": "Protected ID Out of Range",
                            "Message/Signal Name": mes,
                            "Details": f"Protected ID: 0x{pid:02X}",
                            "Expected": "Must be between 0x00 and 0xFF",
                        }
                    )
                    continue

                if not (0x00 <= frame_id <= 0x3F):
                    continue

                pid_bits = [(pid >> i) & 1 for i in range(8)]
                id_bits = pid_bits[:6]
                p0_received = pid_bits[6]
                p1_received = pid_bits[7]

                p0_calculated = id_bits[0] ^ id_bits[1] ^ id_bits[2] ^ id_bits[4]
                p1_calculated = 1 - (id_bits[1] ^ id_bits[3] ^ id_bits[4] ^ id_bits[5])

                calculated_pid = frame_id | (p0_calculated << 6) | (p1_calculated << 7)
                if pid != calculated_pid:
                    all_errors.append(
                        {
                            "Error Type": "Protected ID Calculation Error",
                            "Message/Signal Name": mes,
                            "Details": f"Received: 0x{pid:02X}, Expected: 0x{calculated_pid:02X}",
                            "Expected": f"Frame ID (0x{frame_id:02X}) + P0 ({p0_calculated}) + P1 ({p1_calculated})",
                        }
                    )

                if p0_received != p0_calculated or p1_received != p1_calculated:
                    all_errors.append(
                        {
                            "Error Type": "Protected ID Parity Error",
                            "Message/Signal Name": mes,
                            "Details": f"Received P0,P1: {p0_received}{p1_received}, Expected: {p0_calculated}{p1_calculated}",
                            "Expected": f"P0 = ID0 ⊕ ID1 ⊕ ID2 ⊕ ID4, P1 = ¬(ID1 ⊕ ID3 ⊕ ID4 ⊕ ID5)",
                        }
                    )
        except Exception as e:
            all_errors.append(
                {
                    "Error Type": "Protected ID Parsing Error",
                    "Message/Signal Name": "N/A",
                    "Details": f"Error: {str(e)}",
                    "Expected": "Protected IDs should be valid hex or decimal values",
                }
            )

    # 3. Message ID errors
    try:
        data_frame["Msg ID"] = data_frame["Msg ID"].apply(
            lambda x: (
                int(x, 16) if isinstance(x, str) and x.startswith("0x") else int(x)
            )
        )

        msg_id = dict(zip(data_frame["Msg Name"], data_frame["Msg ID"]))
        msg_type = dict(zip(data_frame["Msg Name"], data_frame["Send Type"]))

        invalid_range = {}
        invalid_unconditional = {}
        invalid_diagnostic = {}
        forbidden_ids = {}

        for mes, id in msg_id.items():
            if not (0x00 <= id <= 0x3D):
                all_errors.append(
                    {
                        "Error Type": "Message ID Out of Range",
                        "Message/Signal Name": mes,
                        "Details": f"ID: 0x{id:02X}",
                        "Expected": "Must be between 0x00 and 0x3D",
                    }
                )

            if id in [0x3E, 0x3F]:
                all_errors.append(
                    {
                        "Error Type": "Forbidden Message ID",
                        "Message/Signal Name": mes,
                        "Details": f"ID: 0x{id:02X}",
                        "Expected": "IDs 0x3E and 0x3F are reserved",
                    }
                )

            frame_type = msg_type.get(mes, "")

            if frame_type == "UF" and not (0x00 <= id <= 0x3B):
                all_errors.append(
                    {
                        "Error Type": "Invalid ID for Unconditional Frame",
                        "Message/Signal Name": mes,
                        "Details": f"ID: 0x{id:02X}, Type: {frame_type}",
                        "Expected": "Unconditional Frames must use IDs 0x00-0x3B",
                    }
                )

            if frame_type == "DF" and not (0x3C <= id <= 0x3D):
                all_errors.append(
                    {
                        "Error Type": "Invalid ID for Diagnostic Frame",
                        "Message/Signal Name": mes,
                        "Details": f"ID: 0x{id:02X}, Type: {frame_type}",
                        "Expected": "Diagnostic Frames must use IDs 0x3C or 0x3D",
                    }
                )
    except Exception as e:
        all_errors.append(
            {
                "Error Type": "Message ID Parsing Error",
                "Message/Signal Name": "N/A",
                "Details": f"Error: {str(e)}",
                "Expected": "Message IDs should be valid hex or decimal values",
            }
        )

    # 4. Message Send Type errors
    msg_send_type = dict(zip(data_frame["Msg Name"], data_frame["Send Type"]))
    valid_types = ["UF", "EF", "SF", "DF"]

    for mes, send_type in msg_send_type.items():
        if send_type not in valid_types:
            all_errors.append(
                {
                    "Error Type": "Invalid Send Type",
                    "Message/Signal Name": mes,
                    "Details": f"Type: {send_type}",
                    "Expected": "Must be UF (Unconditional), EF (Event), SF (Sporadic), or DF (Diagnostic)",
                }
            )

    # 5. Checksum Mode errors
    if "Checksum Mode" in data_frame.columns:
        checksum_modes = dict(zip(data_frame["Msg Name"], data_frame["Checksum Mode"]))
        send_type = dict(zip(data_frame["Msg Name"], data_frame["Send Type"]))

        for mes, mode in checksum_modes.items():
            mode_str = str(mode).strip().lower()
            if mode_str not in ["classic", "enhanced"]:
                all_errors.append(
                    {
                        "Error Type": "Invalid Checksum Mode",
                        "Message/Signal Name": mes,
                        "Details": f"Mode: {mode}",
                        "Expected": "Must be 'Classic' or 'Enhanced'",
                    }
                )

            if send_type[mes] == "DF" and mode_str != "classic":
                all_errors.append(
                    {
                        "Error Type": "Invalid Checksum for Diagnostic Frame",
                        "Message/Signal Name": mes,
                        "Details": f"Mode: {mode}, Type: {send_type[mes]}",
                        "Expected": "Diagnostic Frames must use Classic checksum",
                    }
                )

    # 6. Message Length errors
    msg_len = dict(zip(data_frame["Msg Name"], data_frame["Msg Length"]))

    for mes, length in msg_len.items():
        if length not in [1, 2, 4, 8]:
            all_errors.append(
                {
                    "Error Type": "Invalid Message Length",
                    "Message/Signal Name": mes,
                    "Details": f"Length: {length} bytes",
                    "Expected": "Must be 1, 2, 4, or 8 bytes",
                }
            )

    # 7. Signal Name errors
    invalid_sig_names = []
    too_long_sig_names = []
    sig_name = set(data_frame["Sig Name"].dropna().astype(str))

    for name in sig_name:
        if not re.fullmatch(r"^[A-Za-z0-9_\-]+$", name.strip()):
            all_errors.append(
                {
                    "Error Type": "Invalid Signal Name",
                    "Message/Signal Name": name,
                    "Details": "Contains prohibited characters",
                    "Expected": "Only A-Z, a-z, 0-9, _, - allowed",
                }
            )

        if len(name) > 32:
            all_errors.append(
                {
                    "Error Type": "Too Long Signal Name",
                    "Message/Signal Name": name,
                    "Details": f"Length: {len(name)} characters",
                    "Expected": "Max 32 characters",
                }
            )

    # 8. Signal Description errors
    sig_desc = dict(zip(data_frame["Sig Name"], data_frame["Description"]))

    for sig_name, val in sig_desc.items():
        if pd.isna(val) or str(val).strip() == "":
            all_errors.append(
                {
                    "Error Type": "Missing Signal Description",
                    "Message/Signal Name": sig_name,
                    "Details": "Value is empty",
                    "Expected": "Signal description is required",
                }
            )

    # 9. Response Error errors
    if "Response Error" in data_frame.columns:
        resp_error = dict(zip(data_frame["Sig Name"], data_frame["Response Error"]))

        for sig, val in resp_error.items():
            if pd.notna(val) and str(val).strip() != "" and not str(val).isdigit():
                all_errors.append(
                    {
                        "Error Type": "Invalid Response Error Value",
                        "Message/Signal Name": sig,
                        "Details": f"Value: {val}",
                        "Expected": "Should be numeric or empty",
                    }
                )

    # 10. Signal Positioning errors
    start_byte = dict(zip(data_frame["Sig Name"], data_frame["Start Byte"]))
    start_bit = dict(zip(data_frame["Sig Name"], data_frame["Start Bit"]))
    bit_length = dict(zip(data_frame["Sig Name"], data_frame["Length"]))

    for sig in start_byte.keys():
        byte = start_byte.get(sig)
        bit = start_bit.get(sig)
        length = bit_length.get(sig)

        errors = []

        if byte not in range(0, 8):
            errors.append(f"Invalid start byte: {byte} (must be 0-7)")

        if bit not in range(0, 8):
            errors.append(f"Invalid start bit: {bit} (must be 0-7)")

        if not (1 <= length <= 16):
            errors.append(f"Invalid length: {length} (must be 1-16 bits)")

        if byte is not None and bit is not None and length is not None:
            end_bit = bit + length - 1
            if end_bit > 7:
                errors.append(f"Signal crosses byte boundary (ends at bit {end_bit})")

        if errors:
            all_errors.append(
                {
                    "Error Type": "Signal Positioning Error",
                    "Message/Signal Name": sig,
                    "Details": "; ".join(errors),
                    "Expected": "Signal must fit within one byte (0-7 bits) and be 1-16 bits long",
                }
            )

    # 11. Initial/Invalid Value errors
    init_values = dict(zip(data_frame["Sig Name"], data_frame["Initinal"]))
    invalid_values = dict(zip(data_frame["Sig Name"], data_frame["Invalid"]))

    for sig in init_values.keys():
        init_val = init_values.get(sig)
        inval_val = invalid_values.get(sig)

        errors = []

        if pd.notna(init_val):
            try:
                if isinstance(init_val, str):
                    if init_val.startswith("0x"):
                        int(init_val, 16)
                    else:
                        int(init_val)
            except ValueError:
                errors.append(f"Invalid initial value: {init_val}")

        if pd.notna(inval_val):
            try:
                if isinstance(inval_val, str):
                    if inval_val.startswith("0x"):
                        int(inval_val, 16)
                    else:
                        int(inval_val)
            except ValueError:
                errors.append(f"Invalid invalid value: {inval_val}")

        if errors:
            all_errors.append(
                {
                    "Error Type": "Initial/Invalid Value Error",
                    "Message/Signal Name": sig,
                    "Details": "; ".join(errors),
                    "Expected": "Values should be in hex (0xXX) or decimal format",
                }
            )

    # 12. Min/Max Value errors
    min_vals = dict(zip(data_frame["Sig Name"], data_frame["Min"]))
    max_vals = dict(zip(data_frame["Sig Name"], data_frame["Max"]))

    for sig in min_vals.keys():
        min_val = min_vals.get(sig)
        max_val = max_vals.get(sig)

        if pd.isna(min_val) or pd.isna(max_val):
            continue

        try:
            min_val = float(min_val)
            max_val = float(max_val)
            if min_val > max_val:
                all_errors.append(
                    {
                        "Error Type": "Min/Max Value Mismatch",
                        "Message/Signal Name": sig,
                        "Details": f"Min: {min_val}, Max: {max_val}",
                        "Expected": "Minimum value must be less than or equal to maximum value",
                    }
                )
        except ValueError:
            all_errors.append(
                {
                    "Error Type": "Invalid Min/Max Value Format",
                    "Message/Signal Name": sig,
                    "Details": f"Min: {min_vals.get(sig)}, Max: {max_vals.get(sig)}",
                    "Expected": "Values should be numeric",
                }
            )

    # Create the Excel file if there are errors
    if all_errors:
        error_df = pd.DataFrame(all_errors)
        error_df = error_df[
            ["Error Type", "Message/Signal Name", "Details", "Expected"]
        ]

        with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
            error_df.to_excel(writer, sheet_name="Validation Errors", index=False)
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

        if len(name) > 32:
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
        with st.expander("Names too long (>32 characters)", expanded=True):
            st.warning(f"Found {len(too_long_names)} too long name:")
            st.dataframe(
                pd.DataFrame(
                    {"Name": too_long_names, "Len": [len(n) for n in too_long_names]}
                )
            )

    return False


def validate_protected_id(data_frame: pd.DataFrame) -> bool:
    if "Protected ID" not in data_frame.columns or "Msg ID" not in data_frame.columns:
        st.warning("Protected ID or Msg ID column not found - skipping validation")
        return True

    try:
        data_frame["Protected ID"] = data_frame["Protected ID"].apply(
            lambda x: (
                int(x, 16) if isinstance(x, str) and x.startswith("0x") else int(x)
            )
        )
        data_frame["Msg ID"] = data_frame["Msg ID"].apply(
            lambda x: (
                int(x, 16) if isinstance(x, str) and x.startswith("0x") else int(x)
            )
        )
    except Exception as e:
        st.error(f"Error parsing IDs: {str(e)}")
        return False

    prot_id = dict(zip(data_frame["Msg Name"], data_frame["Protected ID"]))
    msg_id = dict(zip(data_frame["Msg Name"], data_frame["Msg ID"]))

    invalid_range = {}
    invalid_calculation = {}
    invalid_parity = {}

    for mes, pid in prot_id.items():
        frame_id = msg_id.get(mes, -1)

        if not (0x00 <= pid <= 0xFF):
            invalid_range[mes] = f"0x{pid:02X}"
            continue

        if not (0x00 <= frame_id <= 0x3F):
            continue

        pid_bits = [(pid >> i) & 1 for i in range(8)]
        id_bits = pid_bits[:6]
        p0_received = pid_bits[6]
        p1_received = pid_bits[7]

        p0_calculated = id_bits[0] ^ id_bits[1] ^ id_bits[2] ^ id_bits[4]
        p1_calculated = 1 - (id_bits[1] ^ id_bits[3] ^ id_bits[4] ^ id_bits[5])

        calculated_pid = frame_id | (p0_calculated << 6) | (p1_calculated << 7)
        if pid != calculated_pid:
            invalid_calculation[mes] = {
                "Received": f"0x{pid:02X}",
                "Expected": f"0x{calculated_pid:02X}",
                "Frame ID": f"0x{frame_id:02X}",
            }

        if p0_received != p0_calculated or p1_received != p1_calculated:
            invalid_parity[mes] = {
                "Received P0,P1": f"{p0_received}{p1_received}",
                "Expected P0,P1": f"{p0_calculated}{p1_calculated}",
                "Frame ID": f"0x{frame_id:02X}",
            }

    if not any([invalid_range, invalid_calculation, invalid_parity]):
        st.success("All protected IDs are correct and parity bits are valid!")
        return True

    if invalid_range:
        with st.expander(
            "Protected IDs outside valid range (0x00-0xFF)", expanded=True
        ):
            st.error(f"Found {len(invalid_range)} invalid IDs:")
            st.dataframe(
                pd.DataFrame(
                    {
                        "Msg Name": invalid_range.keys(),
                        "Invalid ID": invalid_range.values(),
                    }
                )
            )
            st.info("Protected ID must be 8-bit value (0x00-0xFF)")

    if invalid_calculation:
        with st.expander("Incorrect protected ID calculation", expanded=True):
            st.error(f"Found {len(invalid_calculation)} calculation errors:")
            st.dataframe(pd.DataFrame.from_dict(invalid_calculation, orient="index"))
            st.info(
                "Protected ID should be: Frame ID (bits 0-5) + P0 (bit 6) + P1 (bit 7)"
            )

    if invalid_parity:
        with st.expander("Parity bits mismatch", expanded=True):
            st.error(f"Found {len(invalid_parity)} parity errors:")
            st.dataframe(pd.DataFrame.from_dict(invalid_parity, orient="index"))
            st.markdown(
                """
            **Parity calculation rules:**
            - P0 = ID0 ⊕ ID1 ⊕ ID2 ⊕ ID4
            - P1 = ¬(ID1 ⊕ ID3 ⊕ ID4 ⊕ ID5)
            """
            )

    return False


def validate_messages_id(data_frame: pd.DataFrame) -> bool:
    try:
        data_frame["Msg ID"] = data_frame["Msg ID"].apply(
            lambda x: (
                int(x, 16) if isinstance(x, str) and x.startswith("0x") else int(x)
            )
        )
    except Exception as e:
        st.error(f"Error parsing message IDs: {str(e)}")
        return False

    msg_id = dict(zip(data_frame["Msg Name"], data_frame["Msg ID"]))
    msg_type = dict(zip(data_frame["Msg Name"], data_frame["Send Type"]))

    invalid_range = {}
    invalid_unconditional = {}
    invalid_diagnostic = {}
    forbidden_ids = {}

    for mes, id in msg_id.items():

        if not (0x00 <= id <= 0x3D):
            invalid_range[mes] = f"0x{id:02X}"

        if id in [0x3E, 0x3F]:
            forbidden_ids[mes] = f"0x{id:02X}"

        frame_type = msg_type.get(mes, "")

        if frame_type == "UF" and not (0x00 <= id <= 0x3B):
            invalid_unconditional[mes] = f"0x{id:02X}"

        if frame_type == "DF" and not (0x3C <= id <= 0x3D):
            invalid_diagnostic[mes] = f"0x{id:02X}"

    if not any(
        [invalid_range, invalid_unconditional, invalid_diagnostic, forbidden_ids]
    ):
        st.success("All message IDs are correct!")
        return True

    if invalid_range:
        with st.expander("IDs outside valid range (0x00-0x3D)", expanded=True):
            st.error(f"Found {len(invalid_range)} IDs outside valid range:")
            st.dataframe(
                pd.DataFrame(
                    {
                        "Msg Name": invalid_range.keys(),
                        "Invalid ID": invalid_range.values(),
                    }
                )
            )
            st.info("LIN IDs must be between 0x00 and 0x3D (0-61 decimal)")

    if forbidden_ids:
        with st.expander("Forbidden IDs (0x3E-0x3F)", expanded=True):
            st.error(f"Found {len(forbidden_ids)} forbidden IDs:")
            st.dataframe(
                pd.DataFrame(
                    {
                        "Msg Name": forbidden_ids.keys(),
                        "Forbidden ID": forbidden_ids.values(),
                    }
                )
            )
            st.info("IDs 0x3E and 0x3F (62-63) are reserved and cannot be used")

    if invalid_unconditional:
        with st.expander("Unconditional Frames with invalid IDs", expanded=True):
            st.error(
                f"Found {len(invalid_unconditional)} Unconditional Frames with incorrect IDs:"
            )
            st.dataframe(
                pd.DataFrame(
                    {
                        "Msg Name": invalid_unconditional.keys(),
                        "Invalid ID": invalid_unconditional.values(),
                    }
                )
            )
            st.info("Unconditional Frames must use IDs 0x00-0x3B (0-59)")

    if invalid_diagnostic:
        with st.expander("Diagnostic Frames with invalid IDs", expanded=True):
            st.error(
                f"Found {len(invalid_diagnostic)} Diagnostic Frames with incorrect IDs:"
            )
            st.dataframe(
                pd.DataFrame(
                    {
                        "Msg Name": invalid_diagnostic.keys(),
                        "Invalid ID": invalid_diagnostic.values(),
                    }
                )
            )
            st.info(
                "Diagnostic Frames must use IDs 0x3C (Master Request) or 0x3D (Slave Response)"
            )

    return False


def validate_messages_send_type(data_frame: pd.DataFrame) -> bool:
    msg_send_type = dict(zip(data_frame["Msg Name"], data_frame["Send Type"]))
    invalid_send_type = {}

    valid_types = ["UF", "EF", "SF", "DF"]  # LIN frame types

    for mes, send_type in msg_send_type.items():
        if send_type not in valid_types:
            invalid_send_type[mes] = send_type

    if not invalid_send_type:
        st.success("All messages send types are correct!")
        st.info(
            "Valid types: UF (Unconditional), EF (Event), SF (Sporadic), DF (Diagnostic)"
        )
        return True
    if invalid_send_type:
        with st.expander("Incorrect send types", expanded=True):
            st.error(f"Found {len(invalid_send_type.keys())} incorrect send types:")
            st.dataframe(
                pd.DataFrame(
                    {
                        "Msg Name": invalid_send_type.keys(),
                        "Incorrect type": invalid_send_type.values(),
                    }
                )
            )
            st.info(
                "Send Type should be: UF (Unconditional), EF (Event), SF (Sporadic), DF (Diagnostic)"
            )

    return False


def validate_checksum_mode(data_frame: pd.DataFrame) -> bool:
    if "Checksum Mode" not in data_frame.columns:
        st.error("Checksum Mode column not found")
        return False

    checksum_modes = dict(zip(data_frame["Msg Name"], data_frame["Checksum Mode"]))
    send_type = dict(zip(data_frame["Msg Name"], data_frame["Send Type"]))
    invalid_modes = {}
    invalid_send_modes = {}
    for mes, mode in checksum_modes.items():
        if str(mode).strip().lower() not in ["classic", "enhanced"]:
            invalid_modes[mes] = mode

        if send_type[mes] == "DF" and mode != "classic":
            invalid_send_modes[mes] = mode

    if not invalid_modes and not invalid_send_modes:
        st.success("All checksum modes are correct!")
        return True

    with st.expander("Incorrect checksum modes", expanded=True):
        st.error(f"Found {len(invalid_modes.keys())} incorrect checksum modes:")
        st.dataframe(
            pd.DataFrame(
                {
                    "Msg Name": invalid_modes.keys(),
                    "Incorrect mode": invalid_modes.values(),
                }
            )
        )
        st.info("Checksum mode should be 'Classic' or 'Enhanced' for LIN")

    with st.expander("Incorrect checksum for send type", expanded=True):
        st.error(
            f"Found {len(invalid_send_modes.keys())} incorrect checksum for send type:"
        )
        st.dataframe(
            pd.DataFrame(
                {
                    "Msg Name": invalid_send_modes.keys(),
                    "Incorrect mode": invalid_send_modes.values(),
                }
            )
        )

    return False


def validate_messages_length(data_frame: pd.DataFrame) -> bool:
    msg_len = dict(zip(data_frame["Msg Name"], data_frame["Msg Length"]))
    invalid_len = {}

    for mes, length in msg_len.items():
        if length not in [1, 2, 4, 8]:
            invalid_len[mes] = length

    if not invalid_len:
        st.success("All message lengths are correct (1, 2, 4, or 8 bytes)!")
        return True

    with st.expander("Incorrect message lengths", expanded=True):
        st.error(f"Found {len(invalid_len.keys())} incorrect lengths:")
        st.dataframe(
            pd.DataFrame(
                {
                    "Msg Name": invalid_len.keys(),
                    "Incorrect length": invalid_len.values(),
                }
            )
        )
        st.info("LIN message length must be 1, 2, 4, or 8 bytes")

    return False


def validate_signal_names(data_frame: pd.DataFrame) -> bool:
    invalid_names = []
    too_long_names = []
    sig_name = set(data_frame["Sig Name"].dropna().astype(str))

    for name in sig_name:
        if not re.fullmatch(r"^[A-Za-z0-9_\-]+$", name.strip()):
            invalid_names.append(name)

        if len(name) > 32:
            too_long_names.append(name)

    if not invalid_names and not too_long_names:
        st.success("All signal names are correct!")
        return True

    if invalid_names:
        with st.expander(
            "Incorrect names (contain prohibited characters)", expanded=True
        ):
            st.error(f"Found {len(invalid_names)} incorrect name:")
            st.dataframe(pd.DataFrame({"Incorrect name": invalid_names}))
            st.info("Allowed characters: A-Z, a-z, 0-9, _, -")

    if too_long_names:
        with st.expander("Names too long (>32 characters)", expanded=True):
            st.warning(f"Found {len(too_long_names)} too long name:")
            st.dataframe(
                pd.DataFrame(
                    {"Name": too_long_names, "Len": [len(n) for n in too_long_names]}
                )
            )

    return False


def validate_signal_descriptions(data_frame: pd.DataFrame) -> bool:
    sig_desc = dict(zip(data_frame["Sig Name"], data_frame["Description"]))
    invalid_nan = {}

    for mes, val in sig_desc.items():
        if pd.isna(val) or str(val).strip() == "":
            invalid_nan[mes] = "Missing description"

    if not invalid_nan:
        st.success("All signal descriptions are present!")
        return True

    with st.expander("Missing signal descriptions", expanded=True):
        st.error(f"Found {len(invalid_nan.keys())} missing descriptions:")
        st.dataframe(
            pd.DataFrame(
                {
                    "Signal Name": invalid_nan.keys(),
                    "Status": invalid_nan.values(),
                }
            )
        )

    return False


def validate_response_error(data_frame: pd.DataFrame) -> bool:
    if "Response Error" not in data_frame.columns:
        st.warning("Response Error column not found - skipping validation")
        return True

    resp_error = dict(zip(data_frame["Sig Name"], data_frame["Response Error"]))
    invalid_values = {}

    for sig, val in resp_error.items():
        if pd.notna(val) and str(val).strip() != "" and not str(val).isdigit():
            invalid_values[sig] = val

    if not invalid_values:
        st.success("All response error values are valid!")
        return True

    if invalid_values:
        with st.expander("Valid response error values", expanded=True):
            st.info(f"Found {len(invalid_values.keys())} valid values:")
            st.dataframe(
                pd.DataFrame(
                    {
                        "Signal Name": invalid_values.keys(),
                        "Invalid value": invalid_values.values(),
                    }
                )
            )
            st.info("Response Error should be Yes or empty")

    return False


def validate_signal_positioning(data_frame: pd.DataFrame) -> bool:
    start_byte = dict(zip(data_frame["Sig Name"], data_frame["Start Byte"]))
    start_bit = dict(zip(data_frame["Sig Name"], data_frame["Start Bit"]))
    bit_length = dict(zip(data_frame["Sig Name"], data_frame["Length"]))

    invalid_positions = []

    for sig in start_byte.keys():
        byte = start_byte.get(sig)
        bit = start_bit.get(sig)
        length = bit_length.get(sig)

        errors = []

        if byte not in range(0, 8):
            errors.append(f"Invalid start byte: {byte} (must be 0-7)")

        if bit not in range(0, 64):
            errors.append(f"Invalid start bit: {bit} (must be 0-63)")

        if not (1 <= length <= 16):
            errors.append(f"Invalid length: {length} (must be 1-16 bits)")

        if byte is not None and bit is not None and length is not None:
            end_bit = bit + length - 1
            if end_bit > 63:
                errors.append(f"Signal crosses byte boundary (ends at bit {end_bit})")

        if errors:
            invalid_positions.append(
                {
                    "Signal Name": sig,
                    "Errors": "; ".join(errors),
                    "Start Byte": byte,
                    "Start Bit": bit,
                    "Length": length,
                }
            )

    if not invalid_positions:
        st.success("All signal positions are valid!")
        return True

    with st.expander("Invalid signal positions", expanded=True):
        st.error(f"Found {len(invalid_positions)} invalid signal positions:")
        st.dataframe(pd.DataFrame(invalid_positions))
        st.info("Signal must fit within one byte (0-7 bits) and be 1-16 bits long")

    return False


def validate_start_byte(data_frame: pd.DataFrame) -> bool:
    start_byte = dict(zip(data_frame["Sig Name"], data_frame["Start Byte"]))
    invalid_byte = {}

    for sig, byte in start_byte.items():
        if byte not in range(0, 8):
            invalid_byte[sig] = byte

    if not invalid_byte:
        st.success("All start bytes are correct (0-7)!")
        return True

    with st.expander("Incorrect start bytes", expanded=True):
        st.error(f"Found {len(invalid_byte.keys())} incorrect start bytes:")
        st.dataframe(
            pd.DataFrame(
                {
                    "Signal Name": invalid_byte.keys(),
                    "Incorrect start byte": invalid_byte.values(),
                }
            )
        )
        st.info("Start byte must be between 0 and 7 for LIN")

    return False


def validate_start_bit(data_frame: pd.DataFrame) -> bool:
    start_bit = dict(zip(data_frame["Sig Name"], data_frame["Start Bit"]))
    invalid_bit = {}

    for sig, bit in start_bit.items():
        if bit not in range(0, 64):
            invalid_bit[sig] = bit

    if not invalid_bit:
        st.success("All start bits are correct (0-63)!")
        return True

    with st.expander("Incorrect start bits", expanded=True):
        st.error(f"Found {len(invalid_bit.keys())} incorrect start bits:")
        st.dataframe(
            pd.DataFrame(
                {
                    "Signal Name": invalid_bit.keys(),
                    "Incorrect start bit": invalid_bit.values(),
                }
            )
        )
        st.info("Start bit must be between 0 and 7 for LIN")

    return False


def validate_signal_length(data_frame: pd.DataFrame) -> bool:
    sig_len = dict(zip(data_frame["Sig Name"], data_frame["Length"]))
    invalid_len = {}

    for sig, length in sig_len.items():
        if not (1 <= length <= 16):
            invalid_len[sig] = length

    if not invalid_len:
        st.success("All signal lengths are correct (1-16 bits)!")
        return True

    with st.expander("Incorrect signal lengths", expanded=True):
        st.error(f"Found {len(invalid_len.keys())} incorrect lengths:")
        st.dataframe(
            pd.DataFrame(
                {
                    "Signal Name": invalid_len.keys(),
                    "Incorrect length": invalid_len.values(),
                }
            )
        )
        st.info("Signal length must be between 1 and 16 bits for LIN")

    return False


def validate_initial_invalid_values(data_frame: pd.DataFrame) -> bool:
    init_values = dict(zip(data_frame["Sig Name"], data_frame["Initinal"]))
    invalid_values = dict(zip(data_frame["Sig Name"], data_frame["Invalid"]))

    invalid_entries = []

    for sig in init_values.keys():
        init_val = init_values.get(sig)
        inval_val = invalid_values.get(sig)

        errors = []

        if pd.notna(init_val):
            try:
                if isinstance(init_val, str):
                    if init_val.startswith("0x"):
                        int(init_val, 16)
                    else:
                        int(init_val)
            except ValueError:
                errors.append(f"Invalid initial value: {init_val}")

        if pd.notna(inval_val):
            try:
                if isinstance(inval_val, str):
                    if inval_val.startswith("0x"):
                        int(inval_val, 16)
                    else:
                        int(inval_val)
            except ValueError:
                errors.append(f"Invalid invalid value: {inval_val}")

        if errors:
            invalid_entries.append(
                {
                    "Signal Name": sig,
                    "Errors": "; ".join(errors),
                    "Initial Value": init_val,
                    "Invalid Value": inval_val,
                }
            )

    if not invalid_entries:
        st.success("All initial and invalid values are valid!")
        return True

    with st.expander("Invalid initial/invalid values", expanded=True):
        st.error(f"Found {len(invalid_entries)} invalid values:")
        st.dataframe(pd.DataFrame(invalid_entries))
        st.info("Values should be in hex (0xXX) or decimal format")

    return False


def validate_min_max_values(data_frame: pd.DataFrame) -> bool:
    min_vals = dict(zip(data_frame["Sig Name"], data_frame["Min"]))
    max_vals = dict(zip(data_frame["Sig Name"], data_frame["Max"]))
    invalid_pairs = []

    for sig in min_vals.keys():
        min_val = min_vals.get(sig)
        max_val = max_vals.get(sig)

        if pd.isna(min_val) or pd.isna(max_val):
            continue

        try:
            min_val = float(min_val)
            max_val = float(max_val)
            if min_val > max_val:
                invalid_pairs.append(
                    {"Signal Name": sig, "Min Value": min_val, "Max Value": max_val}
                )
        except ValueError:
            invalid_pairs.append(
                {
                    "Signal Name": sig,
                    "Error": "Invalid numeric format",
                    "Min Value": min_vals.get(sig),
                    "Max Value": max_vals.get(sig),
                }
            )

    if not invalid_pairs:
        st.success("All min/max value pairs are valid!")
        return True

    with st.expander("Invalid min/max value pairs", expanded=True):
        st.error(f"Found {len(invalid_pairs)} invalid min/max pairs:")
        st.dataframe(pd.DataFrame(invalid_pairs))
        st.info("Minimum value must be less than or equal to maximum value")

    return False


def main():
    st.title("🚀LIN Frames Validator")
    uploaded_file = st.file_uploader("Upload matrix file", type=["xlsx", "xls", "xlsm"])

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
                tab71,
                tab72,
                tab8,
                tab9,
                tab10,
                tab11,
                tab12,
                tab13,
            ) = st.tabs(
                [
                    "Message Names",
                    "Protected IDs",
                    "Messages IDs",
                    "Messages Send Type",
                    "Messages Lenght",
                    "Signal Name",
                    "Signal Description",
                    "Response Error",
                    "Signal Positioning",
                    "Start Byte",
                    "Start Bit",
                    "Checksum Mode",
                    "Signal Length",
                    "Initianal-Invalid Value",
                    "Minimum-Maximum",
                ]
            )

            with tab1:
                validate_messages_name(processed_df)

            with tab2:
                validate_protected_id(processed_df)

            with tab3:
                validate_messages_id(processed_df)

            with tab4:
                validate_messages_send_type(processed_df)

            with tab5:
                validate_messages_length(processed_df)

            with tab6:
                validate_signal_names(processed_df)

            with tab7:
                validate_signal_descriptions(processed_df)

            with tab71:
                validate_response_error(processed_df)

            with tab72:
                validate_signal_positioning(processed_df)

            with tab8:
                validate_start_byte(processed_df)

            with tab9:
                validate_start_bit(processed_df)

            with tab10:
                validate_checksum_mode(processed_df)

            with tab11:
                validate_signal_length(processed_df)

            with tab12:
                validate_initial_invalid_values(processed_df)

            with tab13:
                validate_min_max_values(processed_df)

        except Exception as e:
            st.error(f"Error processing file: {str(e)}")
    else:
        st.info("Please upload an Excel file to begin validation")


if __name__ == "__main__":
    main()
