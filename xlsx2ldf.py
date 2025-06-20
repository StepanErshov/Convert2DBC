from ldfparser.schedule import ScheduleTable, LinFrameEntry
from ldfparser.lin import LinVersion
from ldfparser.node import LinNode
from ldfparser.frame import LinUnconditionalFrame
from ldfparser.encoding import PhysicalValue, LogicalValue
from ldfparser import (
    LDF,
    LinMaster,
    LinSlave,
    LinSignal,
    LinNodeComposition,
    LinNodeCompositionConfiguration,
    LinDiagnosticFrame,
    LinSignalEncodingType,
    LinProductId,
    save_ldf,
)
import pandas as pd
from typing import Optional, Dict
import re
import argparse
from streamlit.runtime.uploaded_file_manager import UploadedFile
import os
import datetime


class ValueDescriptionParser:
    @staticmethod
    def parse(desc_str: str) -> Optional[Dict[int, str]]:
        """Convert multi-line hex descriptions to single-line decimal format"""
        if not isinstance(desc_str, str) or not desc_str.strip():
            return None

        descriptions = {}
        try:
            desc_str = " ".join(desc_str.replace("\r", "\n").split())

            range_matches = list(
                re.finditer(
                    r"(0x[0-9a-fA-F]+)\s*~\s*(0x[0-9a-fA-F]+)\s*:\s*([^;]*)", desc_str
                )
            )

            for match in range_matches:
                start = int(match.group(1), 16)
                end = int(match.group(2), 16)
                text = match.group(3).strip()
                descriptions[start] = f"{match.group(1)}~{match.group(2)}, {text}"

            for match in reversed(range_matches):
                desc_str = desc_str[: match.start()].strip() + desc_str[match.end() :]

            parts = re.split(r"(0x[0-9a-fA-F]+)\s*:\s*", desc_str)
            if len(parts) > 1:
                for i in range(1, len(parts), 2):
                    hex_val = parts[i]
                    text = parts[i + 1].split(";")[0].strip()
                    if hex_val and text:
                        try:
                            dec_val = int(hex_val, 16)
                            if dec_val not in descriptions:
                                descriptions[dec_val] = re.sub(
                                    r"[^a-zA-Z0-9_\- ]", "", text
                                )
                        except ValueError:
                            continue

            return dict(sorted(descriptions.items())) if descriptions else None

        except Exception as e:
            print(f"Error parsing value descriptions '{desc_str}': {str(e)}")
            return None


class ExcelToLDFConverter:

    def _get_engine(self, file_path: str) -> str:
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

    def __init__(self, excel_path: str):
        self.excel_path = excel_path
        self.ldf = LDF()
        self.engine = self._get_engine(self.excel_path)

        df = pd.read_excel(
            self.excel_path,
            sheet_name="Matrix",
            keep_default_na=True,
            engine=self.engine,
        )

        self.df_info = pd.read_excel(
            self.excel_path, sheet_name="Info", keep_default_na=True, engine=self.engine
        )

        self.df_schedule = pd.read_excel(
            self.excel_path,
            sheet_name="LIN Schedule",
            keep_default_na=True,
            engine=self.engine,
        )

        self.bus_users = [
            col
            for col in df.columns
            if any(val in ["S", "R"] for val in df[col].dropna().unique())
            and col != "Unit\n单位"
        ]

        self.ldf_version = LinVersion(
            str(self.df_info.iloc[1, 0]).strip(".")[0],
            str(self.df_info.iloc[1, 0]).strip(".")[2],
        )

        self.ldf._protocol_version = self.ldf_version
        self.ldf._language_version = self.ldf_version
        self.ldf._baudrate = self.df_info.iloc[1, 1] * 1000

        self.master = LinMaster(
            name=self.bus_users[0],
            timebase=self.df_info.iloc[1, 2] / 1000,
            jitter=float(self.df_info.iloc[1, 3]) / 1000,
            max_header_length=None,
            response_tolerance=None,
        )

        for i, user in enumerate(self.bus_users[1:]):
            slave = LinSlave(name=user)
            slave.lin_protocol = self.df_info.iloc[6, 2]
            slave.configured_nad = self.df_info.iloc[i + 6, 1]
            slave.product_id = LinProductId(0x0, 0x0, 0)
            slave.p2_min = 0.05
            slave.st_min = 0.0
            slave.n_as_timeout = 1.0
            slave.n_cr_timeout = 1.0

            response_signal_row = df[
                (df[slave.name] == "S") & (df["Response Error"] == "Yes")
            ]

            config_frames_df = df[
                ((df[slave.name] == "S") | (df[slave.name] == "R"))
                & (df["Msg Name\n报文名称"].notna())
            ]

            configurable_frames = {}
            for _, row in config_frames_df.iterrows():
                frame_name = str(row["Msg Name\n报文名称"]).strip()
                frame_id = str(row["Msg ID(hex)\n报文标识符"]).strip()
                configurable_frames[frame_id] = frame_name

            slave.configurable_frames = configurable_frames
            if not response_signal_row.empty:
                signal_name = LinSignal(
                    name=response_signal_row.iloc[0]["Signal Name\n信号名称"],
                    width=response_signal_row.iloc[0]["Bit Length(Bit)\n信号长度"],
                    init_value=int(
                        response_signal_row.iloc[0]["Initial Value(Hex)\n初始值"], 16
                    ),
                )
                slave.response_error = signal_name

            self.ldf._slaves[slave.name] = slave

        self.ldf._master = self.master
        self.ldf._channel = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    def _load_excel_data(self) -> pd.DataFrame:
        df = pd.read_excel(
            self.excel_path,
            sheet_name="Matrix",
            keep_default_na=True,
            engine=self.engine,
        )

        df_schedule = pd.read_excel(
            self.excel_path,
            sheet_name="LIN Schedule",
            keep_default_na=True,
            engine=self.engine,
        )

        senders = []
        receivers = []

        for _, row in df.iterrows():
            row_senders = []
            row_receivers = []

            for bus_users in self.bus_users:
                if bus_users in df.columns:
                    if pd.notna(row[bus_users]) and row[bus_users] == "S":
                        row_senders.append(bus_users)
                    elif pd.notna(row[bus_users]) and row[bus_users] == "R":
                        row_receivers.append(bus_users)

            senders.append(",".join(row_senders) if row_senders else None)
            receivers.append(",".join(row_receivers) if row_receivers else None)

        new_df = pd.DataFrame(
            {
                "Msg name": df["Msg Name\n报文名称"].ffill(),
                "Msg ID": df["Msg ID(hex)\n报文标识符"].ffill(),
                "Protected Id": df["Protected ID (hex)\n保护标识符"].ffill(),
                "Msg Send Type": df["Msg Send Type\n报文发送类型"].ffill(),
                "Checksum mode": df["Checksum mode\n校验方式"].ffill(),
                "Msg Length": df["Msg Length(Byte)\n报文长度"].ffill(),
                "Signal Name": df["Signal Name\n信号名称"],
                "Sig Description": df["Signal Description\n信号描述"],
                "Response Error": df["Response Error"],
                "Start Byte": df["Start Byte\n起始字节"],
                "Start Bit": df["Start Bit\n起始位"],
                "Bit Length": df["Bit Length(Bit)\n信号长度"],
                "Resolution": df["Resolution\n精度"],
                "Offset": df["Offset\n偏移量"],
                "Minimum": df["Signal Min. Value(phys)\n物理最小值"],
                "Maximum": df["Signal Max. Value(phys)\n物理最大值"],
                "Min Hex": df["Signal Min. Value(Hex)\n总线最小值"],
                "Max Hex": df["Signal Max. Value(Hex)\n总线最大值"],
                "Unit": df["Unit\n单位"],
                "Init value": df["Initial Value(Hex)\n初始值"],
                "Invalid value": df["Invalid Value(Hex)\n无效值"],
                "Sig Val Description": df["Signal Value Description(hex)\n信号值描述"],
                "Remarks": df["Remark\n备注"],
                "Senders": senders,
                "Receivers": receivers,
            }
        )

        df_schedule = df_schedule.iloc[1:].reset_index(drop=True)

        new_df = new_df.dropna(subset=["Signal Name"])

        return new_df, df_schedule

    def get_file_info(self, file_name: str):
        file_start = "ATOM_CAN_Matrix_"
        file_start1 = "ATOM_CANFD_Matrix_"
        file_start2 = "ATOM_LIN_Matrix_"
        file_name_only = os.path.splitext(os.path.basename(file_name))[0]

        if file_name_only.startswith(file_start1):
            protocol = "CANFD"
            remaining = file_name_only[len(file_start1) :]
        elif file_name_only.startswith(file_start2):
            protocol = "LIN"
            remaining = file_name_only[len(file_start2) :]
        elif file_name_only.startswith(file_start):
            protocol = "CAN"
            remaining = file_name_only[len(file_start) :]
        else:
            return None

        parts = remaining.split("_")

        domain_name = parts.pop(0)

        version_string = parts.pop(0)
        if version_string.startswith("V"):
            version = version_string[1:]
        else:
            version = version_string

        version_parts = version.split(".")
        if len(version_parts) != 3:
            return None

        if "-" in version_parts[2]:
            version_num, file_date = version_parts[2].split("-")
            version_parts[2] = version_num
        else:
            if parts and re.match(r"^\d{8}$", parts[0]):
                file_date = parts.pop(0)
            else:
                file_date = ""

        version = ".".join(version_parts)

        device_name = "_".join(parts) if parts else ""

        if device_name.startswith("internal"):
            device_name = device_name[8:].lstrip("_")

        return {
            "version": version,
            "date": file_date,
            "device_name": device_name,
            "domain_name": domain_name,
            "protocol": protocol,
        }

    def _create_signals(self, row: pd.Series) -> LinSignal:
        try:
            comment = (
                str(row["Sig Description"]) if pd.notna(row["Sig Description"]) else ""
            )
            comment = re.sub(r"[\u4e00-\u9fff]+", "", comment)
            comment = str.replace(comment, "/", "")
            comment = str.replace(comment, "\n", "")
            unit = str(row["Unit"]) if pd.notna(row["Unit"]) else ""
            unit = str.replace(unit, "nan", "")
            # self.ldf._comments = self.get_file_info(self.excel_path.name)["version"] # if need start in local pc, del .name and all will be work
            value_description = None
            if pd.notna(row["Sig Val Description"]):
                value_description = ValueDescriptionParser.parse(
                    row["Sig Val Description"]
                )

            signal = LinSignal(
                name=str(row["Signal Name"]),
                width=int(row["Bit Length"]),
                init_value=int(row["Init value"], 16),
                comment=comment,
            )

            converters = []

            if value_description:
                for key, val in value_description.items():
                    converters.append(LogicalValue(phy_value=key, info=val))

            converters.append(
                PhysicalValue(
                    phy_min=int(row["Min Hex"], 16),
                    phy_max=int(row["Max Hex"], 16),
                    scale=float(row["Resolution"]),
                    offset=float(row["Offset"]),
                    unit=str(row["Unit"]) if str(row["Unit"]) != "nan" else None,
                )
            )

            encoding_type = LinSignalEncodingType(
                name=signal.name, converters=converters
            )

            encoding_type._signals = [signal]

            signal.encoding_type = encoding_type

            signal.publisher = LinNode(row["Senders"])
            signal.subscribers = [LinNode(row["Receivers"])]

            if signal.encoding_type:
                self.ldf._signal_encoding_types[signal.encoding_type.name] = (
                    signal.encoding_type
                )
                self.ldf._signal_representations[signal] = signal.encoding_type

            return signal

        except Exception as e:
            print(f"Error creating signal {row['Signal Name']}: {str(e)}")
            return None

    def _create_node(self) -> bool:
        try:
            lst_slv = [slave for slave in self.ldf.get_slaves()]
            for slv in self.ldf.get_slaves():
                node_compos = LinNodeComposition(name=slv.name)
                node_attr = LinNodeCompositionConfiguration(name=slv.name)

            node_compos.nodes = lst_slv
            node_attr.compositions = [node_compos]

            return True
        except Exception as e:
            print(f"Error creating node attr: {str(e)}")
            return False

    def _create_schedule_tables(self, df_schedule: pd.DataFrame):
        try:
            schedule_columns = []

            for col_name, column in df_schedule.items():
                first_value = column.iloc[0]

                if pd.isna(first_value) or not isinstance(first_value, str):
                    continue
                else:
                    schedule_columns.append(first_value)

            for schedule_name in schedule_columns:
                matching_cols = [
                    col
                    for col in df_schedule.columns
                    if str(df_schedule[col].iloc[0]) == schedule_name
                ]

                if not matching_cols:
                    print(f"No matching column found for schedule '{schedule_name}'")
                    continue

                slot_col = matching_cols[0]
                msg_col = df_schedule.columns[df_schedule.columns.get_loc(slot_col) + 1]
                delay_col = df_schedule.columns[
                    df_schedule.columns.get_loc(slot_col) + 2
                ]

                entries = []

                for _, row in (
                    df_schedule[[slot_col, msg_col, delay_col]].iloc[2:].iterrows()
                ):
                    if pd.isna(row[msg_col]):
                        continue

                    try:
                        msg_id = int(str(row[msg_col]).strip(), 16)
                        delay = (
                            float(row[delay_col]) if pd.notna(row[delay_col]) else 0.0
                        )

                        all_frames = {
                            **self.ldf._unconditional_frames,
                            **self.ldf._diagnostic_frames,
                        }

                        frame = next(
                            (f for f in all_frames.values() if f.frame_id == msg_id),
                            None,
                        )

                        if frame and frame != None:
                            entry_frame = LinFrameEntry()
                            entry_frame.frame = frame
                            entry_frame.delay = delay / 1000
                            entries.append(entry_frame)

                    except ValueError as e:
                        print(f"Invalid message ID or delay in row {_}: {e}")
                        continue

                if entries:
                    schedule_table = ScheduleTable(name=schedule_name)
                    schedule_table.schedule = entries
                    self.ldf._schedule_tables[schedule_table.name] = schedule_table

        except Exception as e:
            print(f"Error creating schedule tables: {str(e)}")

    def _create_frames(
        self, frame_id: int, frame_name: str, group: pd.DataFrame
    ) -> bool:
        try:
            signals = {}
            sig_ldf = {}
            for i, row in group.iterrows():
                signal = self._create_signals(row)
                if signal:
                    sig_ldf[signal.name] = signal
                    start_bit = int(row["Start Bit"])
                    signals[start_bit] = signal

            self.ldf._signals.update(sig_ldf)

            if not signals:
                return False

            frm_length = int(group["Msg Length"].iloc[0])
            publisher_name = group["Senders"].iloc[0]
            publisher = None

            if publisher_name in self.ldf._slaves:
                publisher = self.ldf._slaves[publisher_name]
            elif publisher_name == self.master.name:
                publisher = self.master

            unconditional_frame = LinUnconditionalFrame(
                frame_id=int(frame_id, 16),
                name=frame_name,
                length=frm_length,
                signals=signals,
                pad_with_zero=True,
            )

            unconditional_frame.publisher = publisher
            self.ldf._unconditional_frames[unconditional_frame.name] = (
                unconditional_frame
            )

            return True
        except Exception as e:
            print(f"Error creating frame {frame_name}: {str(e)}")
            return False

    def _create_default_diagnostic_frames(self):
        try:
            # --- MasterReq ---
            master_signals = {}
            for i in range(8):
                signal_name = f"MasterReqB{i}"
                signal = LinSignal(name=signal_name, width=8, init_value=0)
                signal.publisher = self.master
                master_signals[i * 8] = signal
                self.ldf._diagnostic_signals[signal.name] = signal

            master_frame = LinDiagnosticFrame(
                name="MasterReq",
                frame_id=0x3C,
                length=8,
                signals=master_signals,
                pad_with_zero=True,
            )
            self.ldf._diagnostic_frames[master_frame.name] = master_frame

            # --- SlaveResp ---
            slave_signals = {}
            for i in range(8):
                signal_name = f"SlaveRespB{i}"
                signal = LinSignal(name=signal_name, width=8, init_value=0)
                signal.publisher = None
                slave_signals[i * 8] = signal
                self.ldf._diagnostic_signals[signal.name] = signal

            slave_frame = LinDiagnosticFrame(
                name="SlaveResp",
                frame_id=0x3D,
                length=8,
                signals=slave_signals,
                pad_with_zero=True,
            )
            self.ldf._diagnostic_frames[slave_frame.name] = slave_frame

            return True
        except Exception as e:
            print(f"Error creating default diagnostic frames: {str(e)}")
            return False

    def convert(self, output_path: str = "out.ldf") -> bool:
        try:
            df, df_sch = self._load_excel_data()
            grouped = df.groupby(["Msg ID", "Msg name"])

            if df is None or df.empty:
                print("No valid data found in Matrix sheet")
                return False

            for (frm_id, frm_name), group in grouped:
                self._create_frames(frm_id, frm_name, group)

            self._create_default_diagnostic_frames()

            if not df_sch.empty:
                self._create_schedule_tables(df_sch)
            else:
                print("No schedule information found")

            self._create_node()

            save_ldf(self.ldf, output_path, "./ldf.jinja2")

            print(f"LDF-file successfully created: {output_path}")
            return True
        except Exception as e:
            print(f"Error during conversion: {str(e)}")
            return False


def main():
    parser = argparse.ArgumentParser(description="Convert Excel-files to LDF-files")
    parser.add_argument("--input", required=True, help="Path to Excel-file")
    parser.add_argument("--output", required="output.ldf", help="Output name LDF-file")
    args = parser.parse_args()

    converter = ExcelToLDFConverter(args.input)
    if converter.convert(args.output):
        print("Conversion completed successfully")
    else:
        print("Conversion failed")


if __name__ == "__main__":
    main()
