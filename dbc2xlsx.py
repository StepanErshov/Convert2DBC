import cantools
import cantools.database
import pandas as pd
import openpyxl
from typing import List, Dict
import argparse
import pprint


class DbcRead():
    def __init__(self, dbc_path: str):
        self.dbc_path = dbc_path

    def CreateDB(self):
        db = cantools.database.load_file(self.dbc_path)
        
        result = {}

        for message in db.messages:
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
                "Signals": []}
            
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
                    "Value_description": signal.choices
                }
                result[message.name]["Signals"].append(signal_data)
        bus_users = db.nodes
        
        return result, bus_users

    def _format_value_description(self, choices):
        if not choices:
            return ""
        
        if isinstance(choices, dict):
            lines = []
            for value, desc in sorted(choices.items()):
                lines.append(f"- **0x{int(value):X}**: {desc.strip()}")
            return "\n".join(lines)
        else:
            return str(choices)

    def convert(self, output_path: str = "output.xlsx") -> bool:
        """Main method convert"""
        try:
            lib, ecu = self.CreateDB()
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                messages_data = []
                for msg_name, msg_data in lib.items():
                    frame_format = "StandardCAN" if msg_data["Protocol"] == "CAN" else "StandardCAN_FD"
                    msg_type = "NM" if str(msg_name).startswith("NM_") else "Diag" if str(msg_name).startswith("Diag") else "Normal"
                    ecu_status = {node.name: "" for node in ecu}
                    
                    if msg_data["Senders"]:
                        for sender in msg_data["Senders"]:
                            if sender in ecu_status:
                                ecu_status[sender] = "S"

                    for signal in msg_data["Signals"]:
                        if signal["Receivers"]:
                            for receiver in signal["Receivers"]:
                                if receiver in ecu_status and ecu_status[receiver] != "S":
                                    ecu_status[receiver] = "R"


                    msg_row = {
                        'Msg Name\n报文名称': msg_name,
                        "Msg Type报文类型": msg_type,
                        'Msg ID\n报文标识符': msg_data['Msg_id'],
                        "Msg Send Type\n报文发送类型": msg_data.get("Send_type", ""),
                        'Msg Cycle Time (ms)\n报文周期时间': msg_data['Cycle_time'],
                        "Frame Format\n帧格式": frame_format,
                        "BRS\n传输速率切换标识位": 1 if frame_format == "StandardCAN_FD" else 0, 
                        'Msg Length (Byte)\n报文长度': msg_data['Msg_length'],
                        # 'Bus Name': msg_data['Bus_name'],
                        "Sender": msg_data.get("Senders", ""),
                    }

                    for node, status in ecu_status.items():
                        msg_row[node] = status
                    
                    messages_data.append(msg_row)
                
                pd.DataFrame(messages_data).to_excel(writer, sheet_name='Messages', index=False)
                
                signals_data = []
                for msg_name, msg_data in lib.items():
                    
                    for signal in msg_data['Signals']:
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
                            12: "Event"
                        }

                        gen_sig_send_type = signal['Sgn_Send_Type'].value if hasattr(signal['Sgn_Send_Type'], 'value') else None
        
                        signal_send_type = send_map.get(gen_sig_send_type, "Unknown")
                        sig_row = {
                            "Signal Name\n信号名称": signal['Sgn_name'],
                            "Signal Description\n信号描述": signal["Comment"],
                            "Byte Order\n排列格式(Intel/Motorola)": "Motorola MSB" if signal["Byte_oreder"] == "big_endian" else "Intel",
                            "Start Byte\n起始字节": signal["Start_bit"] // 8,
                            'Start Bit\n起始位': signal['Start_bit'],
                            "Signal Send Type\n信号发送类型": signal_send_type,
                            'Bit Length (Bit)\n信号长度': signal['Sgn_lenght'],
                            "Data Type\n数据类型": "Unsigned" if signal["Is_signed"] == False else "Signed",
                            "Resolution\n精度": signal["Factor"],
                            "Offset\n偏移量": signal['Offset'],
                            'Signal Min. Value (phys)\n物理最小值': signal['Minimum'],
                            'Signal Max. Value (phys)\n物理最大值': signal['Maximum'],
                            "Signal Min. Value (Hex)\n总线最小值": f"0x{int((signal['Minimum'] - signal['Offset']) / signal['Factor']):X}",
                            "Signal Max. Value (Hex)\n总线最大值": f"0x{int((signal['Maximum'] - signal['Offset']) / signal['Factor']):X}",
                            'Initial Value (Hex)\n初始值': f"0x{int(signal['Initinal']):X}" if pd.notna(signal['Initinal']) else "",
                            'Invalid Value(Hex)\n无效值': f"0x{int(signal['Invalid']):X}" if pd.notna(signal["Invalid"]) else "",
                            "Inactive Value (Hex)\n非使能值": "0x0",
                            'Unit\n单位': signal['Unit'],
                            'Receivers': ', '.join(signal['Receivers']) if signal['Receivers'] else '',
                            'Value Descriptions': str(signal['Value_description']) if signal['Value_description'] else ''
                        }

                        signals_data.append(sig_row)
                
                pd.DataFrame(signals_data).to_excel(writer, sheet_name='Signals', index=False)
            
            print(f"Excel file successfully created: {output_path}")
            return True
        except Exception as e:
            print(f"Error during conversion: {str(e)}")
            return False

if __name__ == "__main__":
    converter = DbcRead("C:\\projects\\Convert2DBC\\ATOM_CAN_Matrix_BD_V8.0.0_20250625.dbc")
    converter.convert("output.xlsx")