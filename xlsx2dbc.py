import cantools
import cantools.database
import cantools.database.conversion
import pandas as pd
from cantools.database.can.formats.dbc import DbcSpecifics
from cantools.database.can import Node
import re
import argparse
from typing import Optional, Dict


class ValueDescriptionParser:

    @staticmethod
    def parse(desc_str: str) -> Optional[Dict[int, str]]:
        """Convert multi-line hex descriptions to single-line decimal format"""
        if not isinstance(desc_str, str) or not desc_str.strip():
            return None

        descriptions = {}
        try:
            desc_str = " ".join(desc_str.replace("\r", "\n").split())
            parts = re.split(r"(0x[0-9a-fA-F]+)\s*:\s*", desc_str)
            if len(parts) > 1:
                for i in range(1, len(parts), 2):
                    hex_val = parts[i]
                    text = parts[i + 1].split(";")[0].split("~")[0].strip()
                    text = re.sub(r"[^a-zA-Z0-9_\- ]", "", text)
                    if hex_val and text:
                        try:
                            dec_val = int(hex_val, 16)
                            descriptions[dec_val] = text
                        except ValueError:
                            continue
            else:
                for item in desc_str.split(";"):
                    item = item.strip()
                    if ":" in item:
                        val_part, text = item.split(":", 1)
                        val_part = val_part.strip()
                        text = text.strip()
                        if val_part.startswith("0x"):
                            try:
                                dec_val = int(val_part, 16)
                                descriptions[dec_val] = text
                            except ValueError:
                                continue

            range_matches = re.finditer(
                r"(0x[0-9a-fA-F]+)\s*~\s*(0x[0-9a-fA-F]+)\s*:\s*([^;]+)", desc_str
            )
            for match in range_matches:
                start = int(match.group(1), 16)
                end = int(match.group(2), 16)
                text = match.group(3).strip()
                for val in range(start, end + 1):
                    descriptions[val] = text

            return dict(sorted(descriptions.items())) if descriptions else None

        except Exception as e:
            print(f"Error parsing value descriptions '{desc_str}': {str(e)}")
            return None


class ExcelToDBCConverter:

    def __init__(self, excel_path: str):
        self.excel_path = excel_path
        self.db = cantools.database.can.Database(
            version="", sort_signals=None, strict=False
        )
        self.db.dbc = DbcSpecifics()
        df = pd.read_excel(
            self.excel_path,
            sheet_name="Matrix",
            keep_default_na=True,
            engine="openpyxl",
        )

        self.bus_users = [
            col
            for col in df.columns
            if any(val in ["S", "R"] for val in df[col].dropna().unique())
            and col != "Unit\n单位"
        ]
        self._initialize_nodes()

    def _initialize_nodes(self):
        self.db.nodes.extend([Node(name=bus_name) for bus_name in self.bus_users])

    def _load_excel_data(self) -> pd.DataFrame:
        df = pd.read_excel(
            self.excel_path,
            sheet_name="Matrix",
            keep_default_na=True,
            engine="openpyxl",
        )

        df_history = pd.read_excel(
            self.excel_path,
            sheet_name="History",
            keep_default_na=True,
            engine="openpyxl",
        )

        all_revisions = df_history["Revision Management\n版本管理"].apply(
            lambda x: x.split("版本")[-1] if pd.notna(x) else x
        )

        df_history = df_history.reindex(df.index)

        senders = []
        receivers = []

        for _, row in df.iterrows():
            row_senders = []
            row_receivers = []

            for bus_user in self.bus_users:
                if bus_user in df.columns:
                    if pd.notna(row[bus_user]) and row[bus_user] == "S":
                        row_senders.append(bus_user)
                    elif pd.notna(row[bus_user]) and row[bus_user] == "R":
                        row_receivers.append(bus_user)

            senders.append(",".join(row_senders) if row_senders else "Vector__XXX")
            receivers.append(
                ",".join(row_receivers) if row_receivers else "Vector__XXX"
            )

        new_df = pd.DataFrame(
            {
                "Message ID": df["Msg ID\n报文标识符"],
                "Message Name": df["Msg Name\n报文名称"],
                "Signal Name": df["Signal Name\n信号名称"],
                "Start Byte": df["Start Byte\n起始字节"],
                "Start Bit": df["Start Bit\n起始位"],
                "Length": df["Bit Length (Bit)\n信号长度"],
                "Factor": df["Resolution\n精度"],
                "Offset": df["Offset\n偏移量"],
                "Initinal": df["Initial Value (Hex)\n初始值"],
                "Invalid": df["Invalid Value(Hex)\n无效值"],
                "Min": df["Signal Min. Value (phys)\n物理最小值"],
                "Max": df["Signal Max. Value (phys)\n物理最大值"],
                "Unit": df["Unit\n单位"],
                "Receiver": receivers,
                "Byte Order": df["Byte Order\n排列格式(Intel/Motorola)"],
                "Data Type": df["Data Type\n数据类型"],
                "Cycle Type": df["Msg Cycle Time (ms)\n报文周期时间"],
                "Send Type": df["Msg Send Type\n报文发送类型"],
                "Description": df["Signal Description\n信号描述"],
                "Msg Length": df["Msg Length (Byte)\n报文长度"].ffill(),
                "Signal Value Description": df["Signal Value Description\n信号值描述"],
                "Senders": senders,
            }
        )

        new_df["Unit"] = new_df["Unit"].astype(str)
        new_df["Unit"] = new_df["Unit"].str.replace("Ω", "Ohm", regex=False)
        new_df["Unit"] = new_df["Unit"].str.replace("℃", "degC", regex=False)

        new_df["Message Name"] = new_df["Message Name"].ffill()
        new_df["Message ID"] = new_df["Message ID"].ffill()
        new_df = new_df.dropna(subset=["Signal Name"])
        new_df["Is Signed"] = new_df["Data Type"].str.contains("Signed", na=False)

        return new_df, all_revisions

    def _create_signal(self, row: pd.Series) -> Optional[cantools.database.can.Signal]:
        try:
            comment = str(row["Description"]) if pd.notna(row["Description"]) else ""
            comment = re.sub(r"[\u4e00-\u9fff]+", "", comment)
            comment = str.replace(comment, "/", "")
            byte_order = (
                "big_endian" if row["Byte Order"] == "Motorola MSB" else "little_endian"
            )

            is_float = (
                "Float" in str(row["Data Type"])
                if pd.notna(row["Data Type"])
                else False
            )

            value_descriptions = None
            if pd.notna(row["Signal Value Description"]):
                value_descriptions = ValueDescriptionParser.parse(
                    row["Signal Value Description"]
                )

            receivers = []
            if pd.notna(row["Receiver"]):
                if isinstance(row["Receiver"], str):
                    receivers = row["Receiver"].split(",")
                else:
                    receivers = [str(row["Receiver"])]

            signal = cantools.database.can.Signal(
                name=str(row["Signal Name"]),
                start=int(row["Start Bit"]),
                length=int(row["Length"]),
                byte_order=byte_order,
                is_signed=bool(row["Is Signed"]),
                raw_initial=int(
                    int(row["Initinal"], 16) if int(row["Initinal"], 16) else 0
                ),
                raw_invalid=(
                    int(int(row["Invalid"], 16)) if pd.notna(row["Invalid"]) else None
                ),
                conversion=cantools.database.conversion.LinearConversion(
                    scale=(
                        int(row["Factor"])
                        if pd.notna(row["Factor"]) and row["Factor"].is_integer()
                        else (float(row["Factor"]) if pd.notna(row["Factor"]) else 1.0)
                    ),
                    offset=(
                        int(row["Offset"])
                        if pd.notna(row["Offset"]) and row["Offset"].is_integer()
                        else (float(row["Offset"]) if pd.notna(row["Offset"]) else 0.0)
                    ),
                    is_float=is_float,
                ),
                comment=comment,
                minimum=(
                    int(row["Min"])
                    if pd.notna(row["Min"]) and float(row["Min"]).is_integer()
                    else (float(row["Min"]) if pd.notna(row["Min"]) else None)
                ),
                maximum=(
                    int(row["Max"])
                    if pd.notna(row["Max"]) and float(row["Max"]).is_integer()
                    else (float(row["Max"]) if pd.notna(row["Max"]) else None)
                ),
                unit=str(row["Unit"]) if pd.notna(row["Unit"]) else "",
                receivers=receivers,
                is_multiplexer=False,
            )

            if value_descriptions:
                signal.choices = value_descriptions

            return signal

        except Exception as e:
            print(f"Error creating signal {row['Signal Name']}: {str(e)}")
            return None

    def _create_message(self, msg_id: str, msg_name: str, group: pd.DataFrame) -> bool:
        try:
            frame_id = (
                int(msg_id, 16)
                if isinstance(msg_id, str) and msg_id.startswith("0x")
                else int(msg_id)
            )

            signals = []
            for _, row in group.iterrows():
                signal = self._create_signal(row)
                if signal:
                    signals.append(signal)

            if not signals:
                return False

            # Split senders by comma if it's a string
            senders = []
            if pd.notna(group["Senders"].iloc[0]):
                if isinstance(group["Senders"].iloc[0], str):
                    senders = group["Senders"].iloc[0].split(",")
                else:
                    senders = [str(group["Senders"].iloc[0])]

            message = cantools.database.can.Message(
                frame_id=frame_id,
                name=str(msg_name),
                length=int(group["Msg Length"].iloc[0]),
                signals=signals,
                sort_signals=None,
                cycle_time=(
                    int(group["Cycle Type"].iloc[0])
                    if pd.notna(group["Cycle Type"].iloc[0])
                    else None
                ),
                is_extended_frame=False,
                senders=senders,
                is_fd=True,
                bus_name="SGW-CGW",
                protocol="CANFD",
                send_type=(
                    group["Send Type"].iloc[0]
                    if pd.notna(group["Send Type"].iloc[0])
                    else None
                ),
                comment=None,
            )

            self.db.messages.append(message)
            return True

        except Exception as e:
            print(f"Error creating message {msg_name}: {str(e)}")
            return False

    def convert(self, output_path: str = "output.dbc") -> bool:
        """Main method convert"""
        try:
            df, all_revisions = self._load_excel_data()
            grouped = df.groupby(["Message ID", "Message Name"])

            for (msg_id, msg_name), group in grouped:
                self._create_message(msg_id, msg_name, group)

            revision_lines = [f"Revision:{rev}" for rev in all_revisions]
            global_comment = 'CM_ "' + ",\n".join(revision_lines) + '" ;\n'

            cantools.database.dump_file(self.db, output_path)

            # with open(output_path, "a", encoding="utf-8") as f:
            #     f.write("\n")
            #     f.write(global_comment)

            print(f"DBC-file successfully created: {output_path}")
            return True

        except Exception as e:
            print(f"Error during conversion: {str(e)}")
            return False


def main():
    parser = argparse.ArgumentParser(description="Convert Excel-files to DBC-files")
    parser.add_argument("--input", required=True, help="Path to Excel-file")
    parser.add_argument("--output", default="output.dbc", help="Output name DBC-file")
    args = parser.parse_args()

    converter = ExcelToDBCConverter(args.input)
    if converter.convert(args.output):
        print("Conversion completed successfully")
    else:
        print("Conversion failed")


if __name__ == "__main__":
    main()
