import pandas as pd
from streamlit.runtime.uploaded_file_manager import UploadedFile
from typing import List, Union, Dict
import re
import pprint
import streamlit as st


# pd.set_option('display.max_columns', None)  # Показать все колонки
# pd.set_option('display.max_rows', None)     # Показать все строки
# pd.set_option('display.width', None)       # Автоматическая ширина (без переноса)
# pd.set_option('display.max_colwidth', None)  # Полный текст в ячейках
# pd.set_option('display.expand_frame_repr', False)  # Не переносить колонки на новые строки


def load_xlsx(file_path: str) -> Union[pd.DataFrame, Dict]:
    try:
        if isinstance(file_path, str) or isinstance(file_path, UploadedFile):
            data_frame = pd.read_excel(
                file_path, 
                sheet_name="Matrix",
                keep_default_na=True,
                engine="openpyxl")
            return data_frame
        elif isinstance(file_path, List):
            finally_df = {}
            for file in file_path:
                data_frame = pd.read_excel(
                    file,
                    sheet_name="Matrix",
                    keep_default_na=True,
                    engine="openpyxl"
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
        receivers.append(
            ",".join(row_receivers) if row_receivers else "Vector__XXX"
        )

    new_df = pd.DataFrame({
        "Msg ID": df["Msg ID\n报文标识符"].ffill(),
        "Msg Name": df["Msg Name\n报文名称"].ffill(),
        "Cycle Type": df["Msg Cycle Time (ms)\n报文周期时间"].ffill(),
        "Msg Time Fast": df[
                    "Msg Cycle Time Fast(ms)\n报文发送的快速周期"
                ].ffill(),
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
        "Max": df["Signal Max. Value (phys)\n物理最大值"],
        "Unit": df["Unit\n单位"],
        "Receiver": receivers,
        "Byte Order": df["Byte Order\n排列格式(Intel/Motorola)"],
        "Data Type": df["Data Type\n数据类型"],
        "Description": df["Signal Description\n信号描述"],
        "Signal Value Description": df["Signal Value Description\n信号值描述"],
        "Senders": senders,
        "Signal Send Type": df["Signal Send Type\n信号发送类型"],
        "Inactive value": df["Inactive Value (Hex)\n非使能值"],
    })

    new_df["Unit"] = new_df["Unit"].astype(str)
    new_df["Unit"] = new_df["Unit"].str.replace("Ω", "Ohm", regex=False)
    new_df["Unit"] = new_df["Unit"].str.replace("℃", "degC", regex=False)

    new_df = new_df.dropna(subset=["Sig Name"])
    new_df["Is Signed"] = new_df["Data Type"].str.contains("Signed", na=False)

    return new_df


def validate_messages_name(data_frame: pd.DataFrame) -> bool:
    invalid_names = []
    too_long_names = []
    
    msg_names = set(data_frame["Msg Name"].dropna().astype(str))
    
    for name in msg_names:
        if not re.fullmatch(r'^[A-Za-z0-9_\-]+$', name.strip()):
            invalid_names.append(name)
        
        if len(name) > 64:
            too_long_names.append(name)

    if not invalid_names and not too_long_names:
        st.success("All message titles are correct!")
        return True
    
    if invalid_names:
        with st.expander("Incorrect names (contain prohibited characters)", expanded=True):
            st.error(f"Found {len(invalid_names)} incorrect name:")
            st.dataframe(pd.DataFrame({"Incorrect name": invalid_names}))
            st.info("Allowed characters: A-Z, a-z, 0-9, _, -")
    
    if too_long_names:
        with st.expander("Names too long (>64 characters)", expanded=True):
            st.warning(f"Found {len(too_long_names)} too long name:")
            st.dataframe(pd.DataFrame({
                "Name": too_long_names,
                "Len": [len(n) for n in too_long_names]
            }))
    
    return False

def validate_messages_type(data_frame: pd.DataFrame) -> bool:
    msg_type = dict(zip(data_frame["Msg Name"], data_frame["Msg Type"]))

    invalid_type = {}
    invalid_name = {}

    for key, val in msg_type.items():
        if val not in ['Normal', 'Diag', 'NM']:
            invalid_type[key] = val

        if key.startswith('Diag') and val != 'Diag':
            invalid_name[key] = val

        if key.startswith('NM_') and val != 'NM':
            invalid_name[key] = val
    
    if not invalid_name and not invalid_type:
        st.success("All message types are correct!")
        return True
    
    if invalid_type:
         with st.expander("Incorrect type (Unknown type)", expanded=True):
            st.error(f"Found {len(invalid_type.keys())} incorrect type:")
            st.dataframe(pd.DataFrame({"Mes Name": invalid_type.keys(),
                                       "Incorrect types": invalid_type.values()}))
            st.info("list of allowed values ​​'Normal', 'Diag', 'NM'")

    if invalid_name:
        with st.expander("Incorrect name (Not for this type))", expanded=True):
            st.error(f"Found {len(invalid_name.keys())} incorrect type:")
            st.dataframe(pd.DataFrame({"Incorrect Name": invalid_name.keys(),
                                       "Msg Type": invalid_name.values()}))
            st.info("NM, if Msg Name first 3 characters = 'NM_' and Diag, if Msg Name firsts 4 characters = 'Diag'")

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
        if 0x700 <= id <= 0x7FF and msg_type[mes] != 'Diag':
            invalid_type[mes] = id
        if 0x500 <= id <= 0x5FF and msg_type[mes] != 'NM':
            invalid_type[mes] = id
    
    if not invalid_type and not invalid_id:
        st.success("All message IDs are correct!")
        return True
    
    if invalid_id:
        with st.expander("Incorrect ID (Whether it fits within the range or not))", expanded=True):
            st.error(f"Found {len(invalid_id.keys())} incorrect IDs:")
            st.dataframe(pd.DataFrame({"Msg Name": invalid_id.keys(),
                                       "Incorrect IDs": invalid_id.values()}))
            st.info("Msg ID - Must be in the range 0x001 to 0x7FF (Hex)")
    
    if invalid_type:
        with st.expander("Incorrect ID for Msg Type (Wrong range)", expanded=True):
            st.error(f"Found {len(invalid_type.keys())} incorrect types:")
            st.dataframe(pd.DataFrame({"Msg Name": invalid_type.keys(),
                                       "Incorrect IDs": invalid_type.values(),}))
            st.info("Diag if Message ID is in the range 0x700 to 7FF and NM if Message ID is in the range 0x500 to 5FF")

    return False

def validate_messages_send_type(data_frame: pd.DataFrame) -> bool:
    
    msg_send_type = dict(zip(data_frame["Msg Name"], data_frame["Send Type"]))

    for mes, send_type in msg_send_type.items():
        if 
    
    return False

def main():
    st.title("CAN Messages Validator")
    
    uploaded_file = st.file_uploader("Upload matrix file", type=["xlsx"])
    
    if uploaded_file:
        try:
            df = load_xlsx(uploaded_file)
            processed_df = create_correct_df(df)
            
            st.success("File loaded successfully!")
            
            tab1, tab2, tab3 = st.tabs(["Message Names", "Message Types", "Messages IDs"])
            
            with tab1:
                if st.button("Check Message Names", key="name_check"):
                    validate_messages_name(processed_df)
            
            with tab2:
                if st.button("Check Message Types", key="type_check"):
                    validate_messages_type(processed_df)
            
            with tab3:
                if st.button("Check Message ID", key="id_check"):
                    validate_messages_id(processed_df)

        except Exception as e:
            st.error(f"Error processing file: {str(e)}")
    else:
        st.info("Please upload an Excel file to begin validation")

if __name__ == "__main__":
    main()

# pprint.pprint(load_xlsx("C:\\projects\\Convert2DBC\\ATOM_CANFD_Matrix_SGW-CGW_V5.0.0_20250123.xlsx"))
# pprint.pprint(load_xlsx(["C:\\projects\\Convert2DBC\\ATOM_CANFD_Matrix_SGW-CGW_V5.0.0_20250123.xlsx", "C:\\projects\\Convert2DBC\\ATOM_CAN_Matrix_BD_V1.5.4_0912.xlsx"]))
# x = create_correct_df(load_xlsx("C:\\projects\\Convert2DBC\\ATOM_CANFD_Matrix_SGW-CGW_V5.0.0_20250123.xlsx"))

# pprint.pprint(validate_messages_type(x))