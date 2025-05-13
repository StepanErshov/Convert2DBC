import cantools
import cantools.database
import cantools.database.conversion
import numpy as np
import pandas as pd
from cantools.database.can.formats.dbc import DbcSpecifics, AttributeDefinition, Attribute
from collections import OrderedDict
from cantools.database.can import Node
import re

def calculate_start_bit(start_byte, start_bit, byte_order, length):
    """Calculate start bit considering byte order and signal length."""
    if byte_order == "motorola":
        return (start_byte * 8) + (7 - (start_bit % 8))
    return (start_byte * 8) + start_bit

def calculate_message_length(signals):
    """Calculate minimum message length needed for all signals."""
    max_bit = 0
    for signal in signals:
        end_bit = signal.start + signal.length
        max_bit = max(max_bit, end_bit)
    return (max_bit + 7) // 8

def parse_value_descriptions(desc_str):
    """Convert multi-line hex descriptions to single-line decimal format"""
    if not isinstance(desc_str, str) or not desc_str.strip():
        return None
    
    descriptions = {}
    try:
        desc_str = ' '.join(desc_str.replace('\r', '\n').split())
        
        parts = re.split(r'(0x[0-9a-fA-F]+)\s*:\s*', desc_str)
        if len(parts) > 1:
            for i in range(1, len(parts), 2):
                hex_val = parts[i]
                text = parts[i+1].split(';')[0].split('~')[0].strip()
                text = re.sub(r'[^a-zA-Z0-9_\- ]', '', text)
                if hex_val and text:
                    try:
                        dec_val = int(hex_val, 16)
                        descriptions[dec_val] = text
                    except ValueError:
                        continue
        else:
            for item in desc_str.split(';'):
                item = item.strip()
                if ':' in item:
                    val_part, text = item.split(':', 1)
                    val_part = val_part.strip()
                    text = text.strip()
                    if val_part.startswith('0x'):
                        try:
                            dec_val = int(val_part, 16)
                            descriptions[dec_val] = text
                        except ValueError:
                            continue

        range_matches = re.finditer(r'(0x[0-9a-fA-F]+)\s*~\s*(0x[0-9a-fA-F]+)\s*:\s*([^;]+)', desc_str)
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

df = pd.read_excel(
    "C:\\projects\\ATOM\\convert2dbc\\ATOM_CANFD_Matrix_SGW-CGW_V5.0.0_20250123.xlsx",
    sheet_name="Matrix",
    keep_default_na=True,
    engine="openpyxl"
)
df_history = pd.read_excel(
    "C:\\projects\\ATOM\\convert2dbc\\ATOM_CANFD_Matrix_SGW-CGW_V5.0.0_20250123.xlsx",
    sheet_name="History",
    keep_default_na=True,
    engine="openpyxl"
)

df_history = df_history.reindex(df.index) 

new_df = pd.DataFrame({
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
    "Receiver": np.where(df["SGW_SG"] == "R", "SGW_SG", "CGW_SG"),
    "Byte Order": df["Byte Order\n排列格式(Intel/Motorola)"],
    "Data Type": df["Data Type\n数据类型"],
    "Cycle Type": df["Msg Cycle Time (ms)\n报文周期时间"],
    "Send Type": df["Msg Send Type\n报文发送类型"],
    "Description": df["Signal Description\n信号描述"],
    "Msg Length": df["Msg Length (Byte)\n报文长度"].ffill(),
    "Signal Value Description": df["Signal Value Description\n信号值描述"],
    "Senders": np.where(df["SGW_SG"] == "S", "SGW_SG", "CGW_SG"),
    "Revision": df_history["Revision Management\n版本管理"].apply(lambda x: x.split("版本")[-1] if pd.notna(x) else x)
})

new_df["Message Name"] = new_df["Message Name"].ffill()
new_df["Message ID"] = new_df["Message ID"].ffill()
new_df = new_df.dropna(subset=["Signal Name"])
new_df["Is Signed"] = new_df["Data Type"].str.contains("Signed", na=False)

bus_users = ["SGW_SG", "CGW_SG"]

grouped = new_df.groupby(["Message ID", "Message Name"])

db = cantools.database.can.Database(version="V5.0.0", sort_signals=None)
db.dbc = DbcSpecifics()

nodes = [Node(name=bus_name) for bus_name in bus_users]
db.nodes.extend(nodes)

for (msg_id, msg_name), group in grouped:
    try:
        frame_id = int(msg_id, 16) if isinstance(msg_id, str) and msg_id.startswith("0x") else int(msg_id)
        
        signals = []
        for _, row in group.iterrows():
            try:
                byte_order = "motorola" if row["Byte Order"] == "Motorola MSB" else "intel"
                start_bit = calculate_start_bit(
                    int(row["Start Byte"]),
                    int(row["Start Bit"]),
                    byte_order,
                    int(row["Length"])
                )

                is_float = "Float" in str(row["Data Type"]) if pd.notna(row["Data Type"]) else False
                # print("Signal = ", str(row["Signal Name"]), " Length = ", int(row["Length"]), " Start Bit = ", int(row["Start Bit"]))
                value_descriptions = None
                if pd.notna(row["Signal Value Description"]):
                    value_descriptions = parse_value_descriptions(row["Signal Value Description"])

                signal = cantools.database.can.Signal(
                    name=str(row["Signal Name"]),
                    start=int(row["Start Bit"]),
                    length=int(row["Length"]),
                    byte_order=byte_order,
                    is_signed=bool(row["Is Signed"]),
                    raw_initial=int(int(row["Initinal"], 16)),
                    raw_invalid = int(int(row["Invalid"], 16)) if pd.notna(row["Invalid"]) else None,
                    conversion=cantools.database.conversion.LinearConversion(
                        scale = int(row["Factor"]) if pd.notna(row["Factor"]) and row["Factor"].is_integer() else float(row["Factor"]) if pd.notna(row["Factor"]) else 1.0,
                        offset = int(row["Offset"]) if pd.notna(row["Offset"]) and row["Offset"].is_integer() else float(row["Offset"]) if pd.notna(row["Offset"]) else 0.0,
                        is_float=is_float
                    ),
                    comment=str(row["Description"]) if pd.notna(row["Description"]) else "",
                    minimum=int(row["Min"]) if pd.notna(row["Min"]) and float(row["Min"]).is_integer() else (float(row["Min"]) if pd.notna(row["Min"]) else None),
                    maximum=int(row["Max"]) if pd.notna(row["Max"]) and float(row["Max"]).is_integer() else (float(row["Max"]) if pd.notna(row["Max"]) else None),
                    unit=str(row["Unit"]) if pd.notna(row["Unit"]) else "",
                    receivers=[str(row["Receiver"])] if pd.notna(row["Receiver"]) else [],
                    is_multiplexer=False
                )

                if value_descriptions:
                    signal.choices = value_descriptions
                
                signals.append(signal)
            except Exception as e:
                print(f"Ошибка при создании сигнала {row['Signal Name']}: {str(e)}")
                continue

        used_bits = set()
        overlap_found = False
        for signal in signals:
            start = signal.start
            end = start + signal.length
            for bit in range(start, end):
                if bit in used_bits:
                    print(f"Предупреждение: Перекрытие битов в сообщении {msg_name} (0x{frame_id:X}), сигнал {signal.name} (бит {bit})")
                    overlap_found = True
                    break
                used_bits.add(bit)
            if overlap_found:
                break
        
        if overlap_found:
            print(f"Пропускаю сообщение {msg_name} (0x{frame_id:X}) из-за перекрытия битов")
            continue

        message_length = calculate_message_length(signals)

        message = cantools.database.can.Message(
            frame_id=frame_id,
            name=str(msg_name),
            length=int(row["Msg Length"]),
            signals=signals,
            sort_signals = None,
            cycle_time=int(row["Cycle Type"]) if pd.notna(row["Cycle Type"]) else None,
            is_extended_frame=False,
            senders=[str(row["Senders"])] if pd.notna(row["Senders"]) else [],
            is_fd=True,
            bus_name="SGW-CGW",
            protocol="CANFD",
            send_type=row["Send Type"] if pd.notna(row["Send Type"]) else None,
            comment=None
        )
        
        db.messages.append(message)
    except Exception as e:
        print(f"Ошибка при создании сообщения {msg_name}: {str(e)}")
        continue

all_revisions = new_df['Revision'].dropna()
revision_lines = [f'Revision:{rev}' for rev in all_revisions]

global_comment = 'CM_ "' + ',\n'.join(revision_lines) + '";\n'


output_file = "output.dbc"
try:
    cantools.database.dump_file(db, output_file)
    print(f"DBC-файл успешно создан: {output_file}")

    with open(output_file, "a", encoding='utf-8') as f:
        f.write("\n")
        f.write(global_comment)
    print(f"Глобальный комментарий добавлен в файл: {output_file}")
except Exception as e:
    print(f"Ошибка при сохранении DBC файла: {str(e)}")