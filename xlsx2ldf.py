import ldfparser
from ldfparser.lin import LinVersion
from ldfparser.node import LinNode
from ldfparser.frame import LinFrame, LinUnconditionalFrame, LinSporadicFrame
from ldfparser import LDF, LinMaster, LinSlave, LinFrame, LinSignal, LinNodeComposition, LinNodeCompositionConfiguration, save_ldf
import pandas as pd
from typing import Optional, Dict
import re

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


class ExcelToLDFConverter:
    def __init__(self, excel_path: str):
        self.excel_path = excel_path
        self.ldf = LDF()
        

        df = pd.read_excel(
            self.excel_path,
            sheet_name="Matrix",
            keep_default_na=True,
            engine="openpyxl"
        )

        self.df_info = pd.read_excel(
            self.excel_path, 
            sheet_name="Info",
            keep_default_na=True,
            engine="openpyxl"
        )

        self.df_schedule = pd.read_excel(
            self.excel_path,
            sheet_name="LIN Schedule",
            keep_default_na=True,
            engine="openpyxl"
        )
        
        self.bus_users = [
            col 
            for col in df.columns
            if any(val in ["S", "R"] for val in df[col].dropna().unique())
            and col != "Unit\n单位"
        ]

        self.ldf_version = LinVersion(self.df_info.iloc[1, 0], 
                                      self.df_info.iloc[1, 0])
        self.master = LinMaster(name=self.bus_users[0],
                                timebase=self.df_info.iloc[1, 2],
                                jitter=float(self.df_info.iloc[1, 3]),
                                max_header_length=8,
                                response_tolerance=0.1)
        self.slave = LinSlave(name=self.bus_users[1])
        self.nodes = LinNode(name="ABVGD")
       
        self.unframe = LinUnconditionalFrame(frame_id=1,
                                             name="First_frame",
                                             length=8,
                                             signals={0: LinSignal("DCM_FL_AmbientLightReq", 1, 0),
                                                      1: LinSignal("DCM_FL_InLightWelReq", 1, 0),
                                                      2: LinSignal("DCM_FL_InLightShaWelReq", 1, 0)},
                                                      pad_with_zero=True)
        self.frame = LinFrame(frame_id=1, name="First_frame")
        
    def _load_excel_data(self) -> pd.DataFrame:
        df = pd.read_excel(
            self.excel_path,
            sheet_name="Matrix",
            keep_default_na=True,
            engine="openpyxl"
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
            receivers.append(
                ",".join(row_receivers) if row_receivers else None
            )

        new_df = pd.DataFrame(
            {
                "Message name":df['Msg Name\n报文名称'].ffill(), 
                "Message ID": df['Msg ID(hex)\n报文标识符'].ffill(), 
                "Protected Id": df['Protected ID (hex)\n保护标识符'].ffill(),
                "Msg Send Type": df['Msg Send Type\n报文发送类型'].ffill(), 
                "Checksum mode": df['Checksum mode\n校验方式'].ffill(),
                "Msg Length": df['Msg Length(Byte)\n报文长度'].ffill(), 
                "Signal Name": df['Signal Name\n信号名称'],
                "Sig Description": df['Signal Description\n信号描述'], 
                "Response Error": df['Response Error'], 
                "Start Byte": df['Start Byte\n起始字节'],
                "Start Bit": df['Start Bit\n起始位'], 
                "Bit Length": df['Bit Length(Bit)\n信号长度'], 
                "Resolution": df['Resolution\n精度'],
                "Offset": df['Offset\n偏移量'], 
                "Minimum": df['Signal Min. Value(phys)\n物理最小值'],
                "Maximum": df['Signal Max. Value(phys)\n物理最大值'], 
                "Min Hex": df['Signal Min. Value(Hex)\n总线最小值'],
                "Max Hex": df['Signal Max. Value(Hex)\n总线最大值'], 
                "Unit": df['Unit\n单位'], 
                "Init value": df['Initial Value(Hex)\n初始值'],
                "Invalid value": df['Invalid Value(Hex)\n无效值'], 
                "Sig Val Description": df['Signal Value Description(hex)\n信号值描述'],
                "Remarks": df['Remark\n备注'],
                "Senders": senders,
                "Receivers": receivers
            }
        )
        new_df = new_df.dropna(subset=["Signal Name"])
        self.ldf.frame = self.frame
        save_ldf(self.ldf, "out.ldf", "C:\\projects\\Convert2DBC\\empty.ldf")
        
        return new_df
    
    def _create_signals(self, row: pd.Series):
        try:
            comment = str(row["Sig Description"]) if pd.notna(row["Sig Description"]) else ""
            comment = re.sub(r"[\u4e00-\u9fff]+", "", comment)
            comment = str.replace(comment, "/", "")
            comment = str.replace(comment, "\n", "")
            unit = str(row["Unit"]) if pd.notna(row["Unit"]) else ""
            unit = str.replace(unit, "nan", "")

            value_description = None
            if pd.notna(row["Sig Val Description"]):
                value_description = ValueDescriptionParser.parse(
                    row["Sig Val Description"]
                )

            slaves = []
            if pd.notna(row["Receivers"]):
                if isinstance(row["Receivers"], str):
                    slaves = [row["Receivers"].split(",")]
                    self.slave.append(slaves) 
                else:
                    slaves = [str[row["Receivers"]]]
                    self.slave.append(slaves)
            
            return 

        except Exception as e:
            print(e)
        
        return



print(ExcelToLDFConverter("C:\\projects\\Convert2DBC\\ATOM_LIN_Matrix_DCM_FL-ALM_FL_V4.0.0-20250121.xlsx")._load_excel_data())
