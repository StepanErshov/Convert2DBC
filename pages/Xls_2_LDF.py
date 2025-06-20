import streamlit as st
import pandas as pd
from xlsx2ldf import ExcelToLDFConverter
import os
from datetime import datetime
import re
import logging
from typing import List, Tuple

# st.set_page_config(
#     page_title="Excel to LDF Converter",
#     page_icon="🚗",
#     layout="wide"
# )

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

    return f"{base_name}_V{new_version}_{current_date}.ldf"


def display_validation_results(errors: List[str], warnings: List[str]):
    """Display validation errors and warnings in Streamlit"""
    if errors:
        with st.expander("❌ Validation Errors", expanded=True):
            for error in errors:
                st.markdown(
                    f'<div class="error-box">{error}</div>', unsafe_allow_html=True
                )

    if warnings:
        with st.expander("⚠️ Validation Warnings", expanded=False):
            for warning in warnings:
                st.markdown(
                    f'<div class="warning-box">{warning}</div>', unsafe_allow_html=True
                )


def validate_input_data(uploaded_file) -> Tuple[List[str], List[str]]:
    """Validate the Excel file before conversion"""
    errors = []
    warnings = []

    try:
        required_sheets = ["Matrix", "Info", "LIN Schedule"]
        with pd.ExcelFile(uploaded_file) as xls:
            missing_sheets = [s for s in required_sheets if s not in xls.sheet_names]
            if missing_sheets:
                errors.append(f"Missing required sheets: {', '.join(missing_sheets)}")
                return errors, warnings

        df_info = pd.read_excel(uploaded_file, sheet_name="Info")
        if len(df_info.columns) < 4:
            errors.append(
                "Info sheet must have at least 4 columns with configuration data"
            )
            return errors, warnings

        version_str = str(df_info.iloc[1, 0]).strip(".")
        if not re.match(r"^\d\.\d$", version_str):
            errors.append(
                f"Invalid LIN version format '{version_str}'. Must be in format X.Y"
            )

        try:
            baudrate = float(df_info.iloc[1, 1]) * 1000
            if not (1000 <= baudrate <= 20000):
                warnings.append(
                    f"Baudrate {baudrate} is outside typical LIN range (1-20 kbps)"
                )
        except (ValueError, TypeError):
            errors.append("Invalid baudrate value in Info sheet")

        df_matrix = pd.read_excel(uploaded_file, sheet_name="Matrix")
        required_matrix_columns = [
            "Msg ID(hex)\n报文标识符",
            "Msg Name\n报文名称",
            "Signal Name\n信号名称",
            "Start Byte\n起始字节",
            "Start Bit\n起始位",
            "Bit Length(Bit)\n信号长度",
            "Msg Length(Byte)\n报文长度",
        ]

        missing_columns = [
            col for col in required_matrix_columns if col not in df_matrix.columns
        ]
        if missing_columns:
            errors.append(
                f"Matrix sheet missing required columns: {', '.join(missing_columns)}"
            )
            return errors, warnings

        seen_ids = set()
        for msg_id in df_matrix["Msg ID(hex)\n报文标识符"].dropna().unique():
            try:
                frame_id = int(str(msg_id).strip(), 16)
                if not (0 <= frame_id <= 0x3F):
                    errors.append(
                        f"Invalid LIN frame ID {hex(frame_id)}. Must be 0x00-0x3F"
                    )

                if frame_id in seen_ids:
                    errors.append(f"Duplicate frame ID {hex(frame_id)}")
                seen_ids.add(frame_id)
            except ValueError:
                errors.append(
                    f"Invalid frame ID format '{msg_id}'. Must be hex (e.g. 0x12)"
                )

        for frame_id, group in df_matrix.groupby("Msg ID(hex)\n报文标识符"):
            try:
                frame_id_val = int(str(frame_id).strip(), 16)
                frame_name = group["Msg Name\n报文名称"].iloc[0]

                frame_length = int(float(group["Msg Length(Byte)\n报文长度"].iloc[0]))

                if not (1 <= frame_length <= 8):
                    errors.append(
                        f"Invalid frame length {frame_length} for {frame_name}. "
                        "LIN frames must be 1-8 bytes"
                    )

                used_bits = [False] * (frame_length * 8)

                for _, row in group.iterrows():
                    signal_name = row["Signal Name\n信号名称"]

                    try:
                        start_byte = int(float(row["Start Byte\n起始字节"]))
                        start_bit = int(float(row["Start Bit\n起始位"]))
                        bit_length = int(float(row["Bit Length(Bit)\n信号长度"]))
                    except (ValueError, TypeError) as e:
                        # errors.append(
                        #     f"Invalid numeric value in signal '{signal_name}' in frame {frame_name}"
                        # )
                        continue

                    if start_byte >= frame_length:
                        errors.append(
                            f"Signal '{signal_name}' in frame {frame_name} starts "
                            f"at byte {start_byte} but frame is only {frame_length} bytes"
                        )

                    if start_bit >= 8:
                        errors.append(
                            f"Signal '{signal_name}' in frame {frame_name} has "
                            f"invalid start bit {start_bit}. Must be 0-7"
                        )

                    if bit_length <= 0:
                        errors.append(
                            f"Signal '{signal_name}' in frame {frame_name} has "
                            f"invalid length {bit_length}. Must be > 0"
                        )

                    start_pos = start_byte * 8 + start_bit
                    end_pos = start_pos + bit_length

                    if end_pos > len(used_bits):
                        errors.append(
                            f"Signal '{signal_name}' in frame {frame_name} exceeds "
                            "frame bounds"
                        )
                    else:
                        for i in range(start_pos, end_pos):
                            if used_bits[i]:
                                errors.append(
                                    f"Signal '{signal_name}' in frame {frame_name} "
                                    f"overlaps with another signal at bit position {i}"
                                )
                            used_bits[i] = True

            except Exception as e:
                errors.append(f"Error validating frame {hex(frame_id_val)}: {str(e)}")

        try:
            df_schedule = pd.read_excel(uploaded_file, sheet_name="LIN Schedule")
            if not df_schedule.empty:
                for col in df_schedule.columns:
                    if pd.isna(df_schedule[col].iloc[0]):
                        continue

                    schedule_name = str(df_schedule[col].iloc[0])
                    msg_col = df_schedule.columns[df_schedule.columns.get_loc(col) + 1]
                    delay_col = df_schedule.columns[
                        df_schedule.columns.get_loc(col) + 2
                    ]

                    for idx, row in df_schedule.iloc[2:].iterrows():
                        if pd.isna(row[msg_col]):
                            continue

                        try:
                            msg_id = str(row[msg_col]).strip()
                            if not msg_id:
                                continue

                            frame_id = int(msg_id, 16)
                            if not (0 <= frame_id <= 0x3F):
                                errors.append(
                                    f"Invalid frame ID {hex(frame_id)} in schedule table "
                                    f"'{schedule_name}' at row {idx+3}"
                                )

                            if pd.notna(row[delay_col]):
                                try:
                                    delay = float(row[delay_col])
                                    if delay < 0:
                                        errors.append(
                                            f"Invalid delay {delay} in schedule table "
                                            f"'{schedule_name}' at row {idx+3}. Must be >= 0"
                                        )
                                except ValueError:
                                    errors.append(
                                        f"Invalid delay value '{row[delay_col]}' in schedule "
                                        f"table '{schedule_name}' at row {idx+3}"
                                    )

                        except ValueError:
                            errors.append(
                                f"Invalid message ID '{msg_id}' in schedule table "
                                f"'{schedule_name}' at row {idx+3}"
                            )
        except Exception as e:
            warnings.append(f"Could not validate schedule table: {str(e)}")

    except Exception as e:
        errors.append(f"Error validating input file: {str(e)}")

    return errors, warnings


def main():
    st.markdown(
        '<h1 class="title">📄 Excel to LDF Converter</h1>', unsafe_allow_html=True
    )
    st.markdown(
        "Upload your Excel file containing LIN data to convert it to an LDF file."
    )

    col1, col2 = st.columns([3, 1])

    with col1:
        uploaded_file = st.file_uploader(
            "Choose an Excel file", type=["xls", "xlsx"], key="file_uploader"
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

                errors, warnings = validate_input_data(uploaded_file)
                display_validation_results(errors, warnings)

                # if errors:
                #     st.error("Cannot convert due to validation errors. Please fix the issues in your Excel file.")
                #     return

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
                help="Enter the version number in format X.X.X",
            )

            base_name = generate_base_name(uploaded_file.name)
            default_output_name = generate_default_output_filename(
                uploaded_file.name, new_version
            )
            custom_filename = st.text_input(
                "Output LDF file name",
                value=default_output_name,
                help="You can customize the output file name",
            )

            if not custom_filename.lower().endswith(".ldf"):
                custom_filename += ".ldf"

            st.markdown("**Final LDF file name:**")
            st.code(custom_filename)

            if st.button("Convert to LDF", key="convert_button"):
                with st.spinner("Converting to LDF... Please wait"):
                    try:
                        converter = ExcelToLDFConverter(uploaded_file)
                        if converter.convert(custom_filename):
                            st.markdown(
                                f'<div class="success-box">Conversion completed successfully!</div>',
                                unsafe_allow_html=True,
                            )

                            with open(custom_filename, "rb") as f:
                                bytes_data = f.read()
                                st.download_button(
                                    label="Download LDF File",
                                    data=bytes_data,
                                    file_name=custom_filename,
                                    mime="application/octet-stream",
                                    key="download_button",
                                )
                        else:
                            st.error("Conversion failed. Please check the input data.")

                            if hasattr(converter, "validation_errors"):
                                display_validation_results(
                                    getattr(converter, "validation_errors", []),
                                    getattr(converter, "validation_warnings", []),
                                )

                    except Exception as e:
                        st.error(f"An error occurred during conversion: {str(e)}")
                        import traceback

                        st.code(traceback.format_exc())


if __name__ == "__main__":
    main()
