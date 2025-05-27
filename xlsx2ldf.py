import ldfparser as ldf
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
        self.ldf = ldf.parse_ldf("example.ldf")
        