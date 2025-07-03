import cantools
import cantools.database
import pandas as pd
import openpyxl
from typing import List, Dict
from collections import OrderedDict
import argparse
import pprint


class DbcRead:
    def __init__(self, dbc_path: str):
        self.dbc_path = dbc_path

    def CreateDB(self):
        db = cantools.database.load_file(self.dbc_path)

        result = {}

        for message in db.messages:
            message_attr = message.dbc.attributes
            gen_msg_cycle_time_fast = (
                message_attr.get("GenMsgCycleTimeFast").value
                if message_attr.get("GenMsgCycleTimeFast")
                else None
            )
            gen_msg_nr_rep = (
                message_attr.get("GenMsgNrOfRepetition").value
                if message_attr.get("GenMsgNrOfRepetition")
                else None
            )
            gen_msg_delay = (
                message_attr.get("GenMsgDelayTime").value
                if message_attr.get("GenMsgDelayTime")
                else None
            )
            result[message.name] = {
                "Msg_id": message.frame_id,
                "Msg_length": message.length,
                "Bus_name": message.bus_name,
                "Msg_comment": message.comment,
                "Cycle_time": message.cycle_time,
                "Byte_order": message.header_byte_order,
                "If_fd": message.is_fd,
                "Protocol": message.protocol,
                "Send_type": message.send_type,
                "Senders": message.senders,
                "GenMsgCycleTimeFast": gen_msg_cycle_time_fast,
                "GenMsgNrOfRepetition": gen_msg_nr_rep,
                "GenMsgDelayTime": gen_msg_delay,
                "Signals": [],
            }

            for signal in message.signals:
                signal_data = {
                    "Sgn_name": signal.name,
                    "Start_bit": signal.start,
                    "Sgn_lenght": signal.length,
                    "Byte_oreder": signal.byte_order,
                    "Sgn_Send_Type": signal.dbc.attributes["GenSigSendType"],
                    "Is_signed": signal.is_signed,
                    "Initinal": signal.raw_initial,
                    "Invalid": signal.raw_invalid,
                    "Factor": signal.conversion.scale,
                    "Offset": signal.conversion.offset,
                    "Is_float": signal.conversion.is_float,
                    "Minimum": signal.minimum,
                    "Maximum": signal.maximum,
                    "Unit": signal.unit,
                    "Comment": signal.comment,
                    "Receivers": signal.receivers,
                    "Value_description": signal.choices,
                }
                result[message.name]["Signals"].append(signal_data)
        bus_users = db.nodes

        return result, bus_users

    def _format_value_description(self, choices):
        if not choices or choices == None:
            return ""

        if isinstance(choices, (dict, OrderedDict)):
            lines = []
            for value, desc in choices.items():
                hex_value = f"0x{int(value):X}"
                lines.append(f"{hex_value}: {desc}")
            return "\n".join(lines)
        return str(choices)

    def convert(self, output_path: str = "output.xlsx") -> bool:
        """Main method convert"""
        try:
            lib, ecu = self.CreateDB()

            combined_data = []

            ecu_nodes = [node.name for node in ecu]

            for msg_name, msg_data in lib.items():
                frame_format = (
                    "StandardCAN" if msg_data["Protocol"] == "CAN" else "StandardCAN_FD"
                )
                msg_type = (
                    "NM"
                    if str(msg_name).startswith("NM_")
                    else "Diag" if str(msg_name).startswith("Diag") else "Normal"
                )

                ecu_status = {node: "" for node in ecu_nodes}
                if msg_data["Senders"]:
                    for sender in msg_data["Senders"]:
                        if sender in ecu_status:
                            ecu_status[sender] = "S"

                for signal in msg_data["Signals"]:
                    signal_ecu_status = ecu_status.copy()
                    if signal["Receivers"]:
                        for receiver in signal["Receivers"]:
                            if (
                                receiver in signal_ecu_status
                                and signal_ecu_status[receiver] != "S"
                            ):
                                signal_ecu_status[receiver] = "R"

                    send_map = {
                        0: "Cyclic",
                        1: "OnChange",
                        2: "OnWrite",
                        3: "IfActive",
                        4: "OnChangeWithRepetition",
                        5: "OnWriteWithRepetition",
                        6: "IfActiveWithRepetition",
                        7: "NoSigSendType",
                        8: "OnChangeAndIfActive",
                        9: "OnChangeAndIfActiveWithRepetition",
                        10: "CA",
                        11: "CE",
                        12: "Event",
                    }
                    gen_sig_send_type = (
                        signal["Sgn_Send_Type"].value
                        if hasattr(signal["Sgn_Send_Type"], "value")
                        else None
                    )
                    signal_send_type = send_map.get(gen_sig_send_type, "Unknown")

                    min_hex = (
                        f"0x{int((signal['Minimum'] - signal['Offset']) / signal['Factor']):X}"
                        if signal["Factor"] != 0
                        else "0x0"
                    )
                    max_hex = (
                        f"0x{int((signal['Maximum'] - signal['Offset']) / signal['Factor']):X}"
                        if signal["Factor"] != 0
                        else "0x0"
                    )

                    value_desc = self._format_value_description(
                        signal["Value_description"]
                    )

                    row = {
                        "Msg Name\n报文名称": msg_name,
                        "Msg Type\n报文类型": msg_type,
                        "Msg ID\n报文标识符": msg_data["Msg_id"],
                        "Msg Send Type\n报文发送类型": msg_data.get("Send_type", ""),
                        "Msg Cycle Time (ms)\n报文周期时间": msg_data["Cycle_time"],
                        "Frame Format\n帧格式": frame_format,
                        "BRS\n传输速率切换标识位": (
                            1 if frame_format == "StandardCAN_FD" else 0
                        ),
                        "Msg Length (Byte)\n报文长度": msg_data["Msg_length"],
                        "Signal Name\n信号名称": signal["Sgn_name"],
                        "Signal Description\n信号描述": signal["Comment"],
                        "Byte Order\n排列格式(Intel/Motorola)": (
                            "Motorola MSB"
                            if signal["Byte_oreder"] == "big_endian"
                            else "Intel"
                        ),
                        "Start Byte\n起始字节": signal["Start_bit"] // 8,
                        "Start Bit\n起始位": signal["Start_bit"],
                        "Signal Send Type\n信号发送类型": signal_send_type,
                        "Bit Length (Bit)\n信号长度": signal["Sgn_lenght"],
                        "Data Type\n数据类型": (
                            "Unsigned" if signal["Is_signed"] == False else "Signed"
                        ),
                        "Resolution\n精度": signal["Factor"],
                        "Offset\n偏移量": signal["Offset"],
                        "Signal Min. Value (phys)\n物理最小值": signal["Minimum"],
                        "Signal Max. Value (phys)\n物理最大值": signal["Maximum"],
                        "Signal Min. Value (Hex)\n总线最小值": min_hex,
                        "Signal Max. Value (Hex)\n总线最大值": max_hex,
                        "Initial Value (Hex)\n初始值": (
                            f"0x{int(signal['Initinal']):X}"
                            if pd.notna(signal["Initinal"])
                            else ""
                        ),
                        "Invalid Value(Hex)\n无效值": (
                            f"0x{int(signal['Invalid']):X}"
                            if pd.notna(signal["Invalid"])
                            else ""
                        ),
                        "Inactive Value (Hex)\n非使能值": "0x0",
                        "Unit\n单位": signal["Unit"],
                        "Signal Value Description\n信号值描述": value_desc,
                        "Msg Cycle Time Fast(ms)\n报文发送的快速周期": msg_data[
                            "GenMsgCycleTimeFast"
                        ],
                        "Msg Nr. Of Reption\n报文快速发送的次数": msg_data[
                            "GenMsgNrOfRepetition"
                        ],
                        "Msg Delay Time(ms)\n报文延时时间": msg_data["GenMsgDelayTime"],
                        "Remarks\n备注": "",
                    }

                    for ecu_node in ecu_nodes:
                        row[ecu_node] = signal_ecu_status.get(ecu_node, "")

                    combined_data.append(row)

            df = pd.DataFrame(combined_data)

            df.to_excel(output_path, index=False, engine="openpyxl")

            print(f"Excel file successfully created: {output_path}")
            return True
        except Exception as e:
            print(f"Error during conversion: {str(e)}")
            return False


if __name__ == "__main__":
    converter = DbcRead(
        "C:\\projects\\Convert2DBC\\ATOM_CAN_Matrix_BD_V8.0.0_20250625.dbc"
    )
    converter.convert("output.xlsx")
