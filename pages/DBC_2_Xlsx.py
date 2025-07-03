import streamlit as st
import pandas as pd
from dbc2xlsx import DbcRead
import os
import tempfile
from datetime import datetime
import re
from sqlalchemy import text

conn = st.connection(
    "can_db",
    type="sql",
    dialect="sqlite",
    database="can_database.db",
)

with conn.session as s:
    s.execute(
        text(
            """
        CREATE TABLE IF NOT EXISTS dbc_converted_files (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            original_filename TEXT NOT NULL,
            xlsx_filename TEXT NOT NULL,
            version TEXT NOT NULL,
            conversion_date TEXT NOT NULL,
            file_size INTEGER NOT NULL,
            user_id TEXT NOT NULL
        )
    """
        )
    )
    s.commit()

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

    return f"{base_name}_V{new_version}_{current_date}.xlsx"


def display_errors(errors):
    if errors:
        with st.expander("⚠️ Validation Errors", expanded=True):
            for error in errors:
                st.markdown(
                    f'<div class="error-box">{error}</div>', unsafe_allow_html=True
                )


def display_warnings(warnings):
    if warnings:
        with st.expander("⚠️ Validation Warnings", expanded=False):
            for warning in warnings:
                st.markdown(
                    f'<div class="warning-box">{warning}</div>', unsafe_allow_html=True
                )


def validate_dbc_file(uploaded_file):
    errors = []
    warnings = []

    try:
        temp_dir = tempfile.gettempdir()
        temp_path = os.path.join(temp_dir, uploaded_file.name)

        if not os.path.exists(temp_dir):
            errors.append(f"Temporary directory does not exist: {temp_dir}")
            return errors, warnings

        with open(temp_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        if not os.path.exists(temp_path):
            errors.append(f"Failed to create temporary file: {temp_path}")
            return errors, warnings

        converter = DbcRead(temp_path)
        lib, ecu = converter.CreateDB()

        if not lib:
            errors.append("DBC file contains no messages")

        if not ecu:
            warnings.append("DBC file contains no ECU nodes")

        for msg_name, msg_data in lib.items():
            if not msg_data.get("Signals"):
                warnings.append(f"Message '{msg_name}' contains no signals")

            if not msg_data.get("Senders"):
                warnings.append(f"Message '{msg_name}' has no senders")

        if os.path.exists(temp_path):
            os.remove(temp_path)

    except Exception as e:
        errors.append(f"Error reading the DBC file: {str(e)}")
        try:
            if "temp_path" in locals() and os.path.exists(temp_path):
                os.remove(temp_path)
        except:
            pass

    return errors, warnings


def main():
    st.markdown(
        '<h1 class="title">📊 DBC to Excel Converter</h1>', unsafe_allow_html=True
    )
    st.markdown(
        "Upload your DBC file containing CAN data to convert it to an Excel file."
    )

    col1, col2 = st.columns([3, 1])

    with col1:
        uploaded_file = st.file_uploader(
            "Choose a DBC file", type=["dbc"], key="file_uploader"
        )

        if uploaded_file is not None:
            try:
                temp_dir = tempfile.gettempdir()
                temp_path = os.path.join(temp_dir, uploaded_file.name)

                if not os.path.exists(temp_dir):
                    st.error(f"Temporary directory does not exist: {temp_dir}")
                    return

                with open(temp_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())

                if not os.path.exists(temp_path):
                    st.error(f"Failed to create temporary file: {temp_path}")
                    return

                converter = DbcRead(temp_path)
                lib, ecu = converter.CreateDB()

                preview_data = []
                ecu_nodes = [node.name for node in ecu]

                for msg_name, msg_data in list(lib.items())[:5]:
                    msg_row = {
                        "Message Name": msg_name,
                        "Message ID": f"0x{int(msg_data['Msg_id']):X}",
                        "Message Length": msg_data["Msg_length"],
                        "Cycle Time": msg_data["Cycle_time"],
                        "Signal Name": None,
                        "Signal Description": None,
                        "Start Byte": None,
                        "Start Bit": None,
                        "Bit Length": None,
                        "Data Type": None,
                        "Resolution": None,
                        "Offset": None,
                        "Min Value": None,
                        "Max Value": None,
                        "Unit": None,
                    }

                    for ecu_node in ecu_nodes[:3]:
                        msg_row[ecu_node] = (
                            "S" if ecu_node in msg_data["Senders"] else ""
                        )

                    preview_data.append(msg_row)

                    for signal in msg_data["Signals"][:3]:
                        sig_row = {
                            "Message Name": None,
                            "Message ID": None,
                            "Message Length": None,
                            "Cycle Time": None,
                            "Signal Name": signal["Sgn_name"],
                            "Signal Description": signal["Comment"] or None,
                            "Start Byte": signal["Start_bit"] // 8,
                            "Start Bit": signal["Start_bit"],
                            "Bit Length": signal["Sgn_lenght"],
                            "Data Type": (
                                "Unsigned" if not signal["Is_signed"] else "Signed"
                            ),
                            "Resolution": signal["Factor"],
                            "Offset": signal["Offset"],
                            "Min Value": signal["Minimum"],
                            "Max Value": signal["Maximum"],
                            "Unit": signal["Unit"] or None,
                        }

                        for ecu_node in ecu_nodes[:3]:
                            if ecu_node in msg_data["Senders"]:
                                sig_row[ecu_node] = "S"
                            elif ecu_node in signal["Receivers"]:
                                sig_row[ecu_node] = "R"
                            else:
                                sig_row[ecu_node] = ""

                        preview_data.append(sig_row)

                preview_df = pd.DataFrame(preview_data)
                
                for col in preview_df.columns:
                    if col in ["Message Length", "Cycle Time", "Start Byte", "Start Bit", "Bit Length", "Resolution", "Offset", "Min Value", "Max Value"]:
                        preview_df[col] = pd.to_numeric(preview_df[col], errors='coerce')
                    else:
                        preview_df[col] = preview_df[col].fillna("")
                
                st.subheader("Data Preview")
                st.dataframe(
                    preview_df.style.set_properties(
                        **{
                            "background-color": "#f0f2f6",
                            "color": "#2c3e50",
                            "border": "1px solid #dfe6e9",
                        }
                    )
                )

                errors, warnings = validate_dbc_file(uploaded_file)
                display_errors(errors)
                display_warnings(warnings)

                if errors:
                    st.error(
                        "Cannot convert due to validation errors. Please fix the issues in your DBC file."
                    )
                    return

                if os.path.exists(temp_path):
                    os.remove(temp_path)

            except Exception as e:
                st.error(f"Error reading the DBC file: {str(e)}")
                try:
                    if "temp_path" in locals() and os.path.exists(temp_path):
                        os.remove(temp_path)
                except:
                    pass
                return

    with col2:
        if uploaded_file is not None:
            st.subheader("Output Settings")

            version, _ = extract_version_date(uploaded_file.name)
            default_version = version if version else "1.0.0"

            new_version = st.text_input(
                "Excel Version",
                value=default_version,
                help="Enter the version number in format X.X.X",
            )

            base_name = generate_base_name(uploaded_file.name)
            default_output_name = generate_default_output_filename(
                uploaded_file.name, new_version
            )

            custom_filename = st.text_input(
                "Output Excel file name",
                value=default_output_name,
                help="You can customize the output file name",
            )

            if not custom_filename.lower().endswith(".xlsx"):
                custom_filename += ".xlsx"

            st.markdown("**Final Excel file name:**")
            st.code(custom_filename)

            if st.button("Convert to Excel", key="convert_button"):
                with st.spinner("Converting... Please wait"):
                    try:
                        temp_path = os.path.join(
                            tempfile.gettempdir(), uploaded_file.name
                        )
                        with open(temp_path, "wb") as f:
                            f.write(uploaded_file.getbuffer())

                        converter = DbcRead(temp_path)
                        success = converter.convert(custom_filename)
                        
                        st.info(f"Conversion result: {success}")
                        st.info(f"Current directory: {os.getcwd()}")
                        st.info(f"File exists: {os.path.exists(custom_filename)}")

                        if success:
                            st.markdown(
                                f'<div class="success-box">Conversion completed successfully!</div>',
                                unsafe_allow_html=True,
                            )

                            # Проверяем, что файл действительно создался
                            if os.path.exists(custom_filename):
                                file_size = os.path.getsize(custom_filename)
                                current_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                                with conn.session as s:
                                    s.execute(
                                        text(
                                            """
                                            INSERT INTO dbc_converted_files (
                                                original_filename, 
                                                xlsx_filename, 
                                                version, 
                                                conversion_date, 
                                                file_size,
                                                user_id
                                            ) VALUES (:original, :xlsx, :version, :date, :size, :user)
                                        """
                                        ),
                                        {
                                            "original": uploaded_file.name,
                                            "xlsx": custom_filename,
                                            "version": new_version,
                                            "date": current_date,
                                            "size": file_size,
                                            "user": st.session_state.get(
                                                "keycloak", {}
                                            ).get("username", "unknown"),
                                        },
                                    )
                                    s.commit()

                                # Читаем файл и создаем кнопку скачивания
                                try:
                                    with open(custom_filename, "rb") as f:
                                        bytes_data = f.read()
                                    
                                    if bytes_data:
                                        st.download_button(
                                            label="Download Excel File",
                                            data=bytes_data,
                                            file_name=custom_filename,
                                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                            key="download_button",
                                        )
                                    else:
                                        st.error("Generated file is empty")
                                except Exception as read_error:
                                    st.error(f"Error reading generated file: {str(read_error)}")
                            else:
                                st.error(f"File was not created: {custom_filename}")

                        if os.path.exists(temp_path):
                            os.remove(temp_path)

                        st.subheader("Conversion History")
                        with conn.session as s:
                            result = s.execute(
                                text(
                                    """
                                SELECT original_filename, xlsx_filename, version, conversion_date, file_size
                                FROM dbc_converted_files
                                ORDER BY conversion_date DESC
                                LIMIT 10
                            """
                                )
                            )
                            history_df = pd.DataFrame(
                                result.fetchall(), columns=result.keys()
                            )

                        if not history_df.empty:
                            history_df["file_size"] = history_df["file_size"].apply(
                                lambda x: (
                                    f"{x/1024:.2f} KB"
                                    if x < 1024 * 1024
                                    else f"{x/(1024*1024):.2f} MB"
                                )
                            )
                            st.dataframe(history_df)

                    except Exception as e:
                        st.error(f"An error occurred: {str(e)}")
                        st.error(f"Full error: {repr(e)}")


conn.session.close()

if __name__ == "__main__":
    main()
