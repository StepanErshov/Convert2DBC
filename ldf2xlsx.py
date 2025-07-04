import canmatrix.formats
import pandas as pd
import argparse
import pprint
import re
from typing import List, Dict, Any
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from pydantic import BaseModel, FilePath, ValidationError
from typing import Dict, Union, List


MATRIX_COLUMNS = [
    "Msg Name\n报文名称",
    "Msg ID(hex)\n报文标识符",
    "Protected ID (hex)\n保护标识符",
    "Msg Send Type\n报文发送类型",
    "Checksum mode\n校验方式",
    "Msg Length(Byte)\n报文长度",
    "Signal Name\n信号名称",
    "Signal Description\n信号描述",
    "Response Error",
    "Start Byte\n起始字节",
    "Start Bit\n起始位",
    "Bit Length(Bit)\n信号长度",
    "Resolution\n精度",
    "Offset\n偏移量",
    "Signal Min. Value(phys)\n物理最小值",
    "Signal Max. Value(phys)\n物理最大值",
    "Signal Min. Value(Hex)\n总线最小值",
    "Signal Max. Value(Hex)\n总线最大值",
    "Unit\n单位",
    "Initial Value(Hex)\n初始值",
    "Invalid Value(Hex)\n无效值",
    "Signal Value Description(hex)\n信号值描述",
    "Remark\n备注",
]

INFO_PARAMS = [
    ("LIN Protocol Version", "LIN 协议版本"),
    ("LIN Baudrate (kbit/s)", "LIN 波特率"),
    ("Time Base (ms)", "基时"),
    ("Jitter (ms)", "偏移"),
]
INFO_ECU_COLUMNS = [
    "ECU Name\n节点名称",
    "NAD (hex)\n节点地址",
    "LIN Protocol Version\nLIN 协议版本",
]
SCHEDULE_COLUMNS = [
    "Slot ID\n时隙编号",
    "Transmitted Msg ID (hex)\n发送报文标识符",
    "Delay (ms)\n时隙",
]


class FilePathModel(BaseModel):
    f: FilePath


def read_file_ldf(file_path: str) -> pd.DataFrame:
    try:
        path_model = FilePathModel(f=file_path)
        validated_path = path_model.f

        data = {}

        with open(file=validated_path, mode="r+", encoding="utf-8") as file:
            data = file.read()

        return data

    except ValidationError as e:
        print(f"Invalid file path: {e}")
        return pd.DataFrame(columns=MATRIX_COLUMNS)
    except Exception as e:
        print(f"Error reading LDF file: {e}")
        return pd.DataFrame(columns=MATRIX_COLUMNS)


def extract_info(data: str) -> Dict[str, Union[str, float]]:
    try:
        lines = data.split("\n")
        for i, line in enumerate(lines):
            if "LIN_description_file;" in line:
                info_lines = lines[i + 1 : i + 5]
                info_dict = {}
                for line in info_lines:
                    if "=" in line:
                        key, value = line.split("=", 1)
                        key = key.strip()
                        value = value.strip().rstrip(";").replace('"', "")
                        info_dict[key] = float(value) if " " not in value else value
                return info_dict
        return {}

    except Exception as e:
        print(f"Error extracting file info: {e}")
        return {}


def extract_nodes(data: str) -> Dict[str, Union[str, float, List[str]]]:
    try:
        lines = [line.strip() for line in data.split("\n")]
        nodes_dict = {
            "Master": {"name": "", "parameters": [], "time_units": "ms"},
            "Slaves": [],
        }

        in_nodes_section = False

        for line in lines:
            if "Nodes {" in line:
                in_nodes_section = True
                continue

            if in_nodes_section:
                if "}" in line:
                    break

                if "Master:" in line:
                    master_parts = (
                        line.split("Master:")[1].strip().rstrip(";").split(",")
                    )
                    nodes_dict["Master"]["name"] = master_parts[0].strip()

                    for param in master_parts[1:]:
                        param = param.strip()
                        if "ms" in param:
                            value = param.replace("ms", "").strip()
                            try:
                                nodes_dict["Master"]["parameters"].append(float(value))
                            except ValueError:
                                pass

                elif "Slaves:" in line:
                    slaves_part = line.split("Slaves:")[1].strip().rstrip(";")
                    nodes_dict["Slaves"] = [s.strip() for s in slaves_part.split(",")]

        return nodes_dict

    except Exception as e:
        print(f"Error extracting nodes: {e}")
        return {
            "Master": {"name": "", "parameters": [], "time_units": "ms"},
            "Slaves": [],
        }


def extract_signals(data: str) -> Dict[str, Dict[str, Union[str, int, List[str]]]]:
    try:
        lines = [line.strip() for line in data.split("\n")]
        signals_dict = {}
        in_signals_section = False

        for line in lines:
            if "Signals {" in line:
                in_signals_section = True
                continue

            if in_signals_section:
                if "}" in line:
                    break

                if not line or line.startswith("//"):
                    continue

                if "//" in line:
                    signal_part, comment = line.split("//", 1)
                    comment = comment.strip()
                else:
                    signal_part = line
                    comment = ""

                signal_part = signal_part.strip().rstrip(";")

                parts = [p.strip() for p in signal_part.split(":")]
                if len(parts) != 2:
                    continue

                signal_name = parts[0]
                signal_params = [p.strip() for p in parts[1].split(",")]

                if len(signal_params) < 4:
                    continue

                try:
                    size = int(signal_params[0])
                    init_value = int(signal_params[1])
                    publishers = [signal_params[2]]
                    subscribers = signal_params[3:]

                    signals_dict[signal_name] = {
                        "size": size,
                        "init_value": init_value,
                        "publishers": publishers,
                        "subscribers": subscribers,
                        "comment": comment,
                    }
                except (ValueError, IndexError) as e:
                    print(f"Error parsing signal {signal_name}: {e}")
                    continue

        return signals_dict

    except Exception as e:
        print(f"Error extracting signals: {e}")
        return {}


def extract_frames(
    data: str,
) -> Dict[str, Dict[str, Union[int, str, List[Dict[str, Union[str, int]]]]]]:
    try:
        lines = [line.strip() for line in data.split("\n")]
        frames_dict = {}
        in_frames_section = False
        current_frame = None

        for line in lines:
            if "Frames {" in line:
                in_frames_section = True
                continue

            if in_frames_section:
                if "}" in line:
                    break

                if not line:
                    continue

                if ":" in line and "{" in line:
                    frame_part = line.split("{")[0].strip()
                    frame_name, frame_params = frame_part.split(":", 1)
                    frame_name = frame_name.strip()

                    params = [p.strip() for p in frame_params.split(",")]
                    if len(params) >= 3:
                        frame_id = int(params[0])
                        publisher = params[1]
                        size = int(params[2])

                        frames_dict[frame_name] = {
                            "frame_id": frame_id,
                            "publisher": publisher,
                            "lenght": size,
                            "signals": [],
                        }
                        current_frame = frame_name
                    continue

                if current_frame and "," in line and ";" in line:
                    signal_part = line.split(";")[0].strip()
                    signal_name, start_bit = signal_part.split(",")
                    signal_name = signal_name.strip()
                    start_bit = int(start_bit.strip())

                    frames_dict[current_frame]["signals"].append(
                        {"signal_name": signal_name, "start_bit": start_bit}
                    )

        return frames_dict

    except Exception as e:
        print(f"Error extracting frames: {e}")
        return {}


def extract_node_attributes(
    data: str,
) -> Dict[str, Dict[str, Union[str, int, float, List[str]]]]:
    try:
        lines = [line.strip() for line in data.split("\n")]
        node_attrs = {}
        brace_count = 0
        in_node_section = False
        in_configurable_frames = False
        current_node = None

        for line in lines:
            brace_count += line.count("{") - line.count("}")

            if "Node_attributes {" in line:
                in_node_section = True
                continue

            if in_node_section:
                if brace_count <= 0 and line.strip() == "}":
                    break

                if line.endswith("{") and not line.startswith("configurable_frames"):
                    node_name = line.split("{")[0].strip()
                    node_attrs[node_name] = {
                        "LIN_protocol": "",
                        "configured_NAD": 0,
                        "product_id": [0, 0, 0],
                        "response_error": "",
                        "P2_min": 0.0,
                        "ST_min": 0.0,
                        "N_As_timeout": 0.0,
                        "N_Cr_timeout": 0.0,
                        "configurable_frames": [],
                    }
                    current_node = node_name
                    continue

                if not current_node:
                    continue

                if "configurable_frames {" in line:
                    in_configurable_frames = True
                    continue

                if in_configurable_frames and line.strip() == "}":
                    in_configurable_frames = False
                    continue

                if in_configurable_frames and current_node:
                    frame = line.strip().rstrip(";").strip()
                    if frame and frame != "}":
                        node_attrs[current_node]["configurable_frames"].append(frame)
                    continue

                if not in_configurable_frames and "=" in line:
                    attr, value = line.split("=", 1)
                    attr = attr.strip()
                    value = value.strip().rstrip(";").strip()

                    if attr == "LIN_protocol":
                        node_attrs[current_node][attr] = value.strip('"')
                    elif attr == "configured_NAD":
                        node_attrs[current_node][attr] = (
                            int(value, 16) if "0x" in value else int(value)
                        )
                    elif attr == "product_id":
                        ids = [
                            int(x.strip(), 16) if "0x" in x else int(x.strip())
                            for x in value.split(",")
                        ]
                        node_attrs[current_node][attr] = ids
                    elif attr in ["P2_min", "ST_min", "N_As_timeout", "N_Cr_timeout"]:
                        num = float(value.split()[0])
                        node_attrs[current_node][attr] = num
                    else:
                        node_attrs[current_node][attr] = value

        return node_attrs

    except Exception as e:
        print(f"Error extracting node attributes: {e}")
        return {}


from typing import Dict, List, Union


def extract_schedule_tables(data: str) -> Dict[str, List[Dict[str, Union[str, int]]]]:
    try:
        lines = [line.strip() for line in data.split("\n")]
        schedules = {}
        in_schedule_section = False
        current_schedule = None
        brace_count = 0

        for line in lines:
            brace_count += line.count("{") - line.count("}")

            if "Schedule_tables {" in line:
                in_schedule_section = True
                continue

            if in_schedule_section:
                if brace_count <= 0 and line.strip() == "}":
                    break

                if line.endswith("{"):
                    schedule_name = line.split("{")[0].strip()
                    schedules[schedule_name] = []
                    current_schedule = schedule_name
                    continue

                if not current_schedule:
                    continue

                if "delay" in line and current_schedule:
                    parts = line.split()
                    if len(parts) >= 4:
                        frame = parts[0]
                        delay = int(parts[2])
                        schedules[current_schedule].append(
                            {"frame": frame, "delay": delay, "unit": "ms"}
                        )

        return schedules

    except Exception as e:
        print(f"Error extracting schedule tables: {e}")
        return {}


def extract_signal_encoding_types(
    data: str,
) -> Dict[str, Dict[str, List[Dict[str, Union[str, int]]]]]:
    try:
        lines = [line.strip() for line in data.split("\n")]
        encodings = {}
        in_encoding_section = False
        current_signal = None
        brace_count = 0

        for line in lines:
            brace_count += line.count("{") - line.count("}")

            if "Signal_encoding_types {" in line:
                in_encoding_section = True
                continue

            if in_encoding_section:
                if brace_count <= 0 and line.strip() == "}":
                    break

                if line.endswith("{"):
                    signal_name = line.split("{")[0].strip()
                    encodings[signal_name] = {
                        "logical_values": [],
                        "physical_values": {},
                    }
                    current_signal = signal_name
                    continue

                if not current_signal:
                    continue

                if "logical_value," in line:
                    parts = [p.strip().strip('"') for p in line.split(",")]
                    if len(parts) >= 3:
                        value = int(parts[1])
                        description = parts[2].rstrip(";").strip()
                        encodings[current_signal]["logical_values"].append(
                            {"value": value, "description": description}
                        )

                elif "physical_value," in line:
                    parts = [p.strip().strip('"') for p in line.split(",")]
                    if len(parts) >= 6:
                        encodings[current_signal]["physical_values"] = {
                            "min": int(parts[1]),
                            "max": int(parts[2]),
                            "scale": float(parts[3]),
                            "offset": float(parts[4]),
                            "unit": parts[5].rstrip(";").strip(),
                        }

        return encodings

    except Exception as e:
        print(f"Error extracting signal encoding types: {e}")
        return {}


def ldf_dicts_to_xlsx(
    info_dict: Dict[str, Any],
    master_slave_dict: Dict[str, Any],
    signals_dict: Dict[str, Any],
    frames_dict: Dict[str, Any],
    node_attrs_dict: Dict[str, Any],
    schedules_dict: Dict[str, Any],
    signal_values_dict: Dict[str, Any],
    output_path: str = "output_ldf.xlsx",
):
    info_columns = [
        "LIN Protocol Version\nLIN协议版本",
        "LIN Baudrate (kbit/s)",
        "Time Base  (ms) \n基时",
        "Jitter (ms)",
    ]
    info_row = [
        info_dict.get("LIN_protocol_version", ""),
        info_dict.get("LIN_speed", ""),
        master_slave_dict.get("Master", {}).get("parameters", [None, None])[0],
        master_slave_dict.get("Master", {}).get("parameters", [None, None])[1],
    ]
    df_info = pd.DataFrame([info_row], columns=info_columns)

    matrix_columns = [
        "Msg Name\n报文名称",
        "Msg ID(hex)\n报文标识符",
        "Protected ID (hex)\n保护标识符",
        "Msg Send Type\n报文发送类型",
        "Checksum mode\n校验方式",
        "Msg Length(Byte)\n报文长度",
        "Signal Name\n信号名称",
        "Signal Description\n信号描述",
        "Response Error",
        "Start Byte\n起始字节",
        "Start Bit\n起始位",
        "Bit Length(Bit)\n信号长度",
        "Resolution\n精度",
        "Offset\n偏移量",
        "Signal Min. Value(phys)\n物理最小值",
        "Signal Max. Value(phys)\n物理最大值",
        "Signal Min. Value(Hex)\n总线最小值",
        "Signal Max. Value(Hex)\n总线最大值",
        "Unit\n单位",
        "Initial Value(Hex)\n初始值",
        "Invalid Value(Hex)\n无效值",
        "Signal Value Description(hex)\n信号值描述",
        "Remark\n备注",
        "BCM",
        "ALM1",
        "ALM2",
    ]
    matrix_rows = []
    for frame_name, frame in frames_dict.items():
        msg_id = frame.get("frame_id", "")
        msg_len = frame.get("lenght", "")
        publisher = frame.get("publisher", "")
        for sig in frame.get("signals", []):
            sig_name = sig["signal_name"]
            sig_props = signals_dict.get(sig_name, {})
            value_desc = ""
            if sig_name in signal_values_dict:
                lv = signal_values_dict[sig_name].get("logical_values", [])
                value_desc = "; ".join(f"{v['value']}={v['description']}" for v in lv)
            bcm = (
                "S"
                if publisher == "BCM"
                else ("R" if "BCM" in sig_props.get("subscribers", []) else "")
            )
            alm1 = (
                "S"
                if publisher == "ALM1"
                else ("R" if "ALM1" in sig_props.get("subscribers", []) else "")
            )
            alm2 = (
                "S"
                if publisher == "ALM2"
                else ("R" if "ALM2" in sig_props.get("subscribers", []) else "")
            )
            matrix_rows.append(
                [
                    frame_name,
                    f"0x{msg_id:X}" if isinstance(msg_id, int) else msg_id,
                    "",
                    "UF",
                    "Enhanced",
                    msg_len,
                    sig_name,
                    sig_props.get("comment", ""),
                    (
                        sig_props.get("publishers", [""])[0]
                        if "response_error" in sig_props
                        else ""
                    ),
                    "",
                    sig.get("start_bit", ""),
                    sig_props.get("size", ""),
                    (
                        sig_props.get("physical_values", {}).get("scale", 1.0)
                        if "physical_values" in sig_props
                        else 1.0
                    ),
                    (
                        sig_props.get("physical_values", {}).get("offset", 0.0)
                        if "physical_values" in sig_props
                        else 0.0
                    ),
                    sig_props.get("physical_values", {}).get("min", ""),
                    sig_props.get("physical_values", {}).get("max", ""),
                    sig_props.get("physical_values", {}).get("min", ""),
                    sig_props.get("physical_values", {}).get("max", ""),
                    sig_props.get("physical_values", {}).get("unit", ""),
                    sig_props.get("init_value", ""),
                    "",
                    value_desc,
                    "",
                    bcm,
                    alm1,
                    alm2,
                ]
            )
    df_matrix = pd.DataFrame(matrix_rows, columns=matrix_columns)

    schedule_columns = ["Schedule Name", "Frame", "Delay", "Unit"]
    schedule_rows = []
    for sched_name, sched_list in schedules_dict.items():
        for item in sched_list:
            schedule_rows.append(
                [
                    sched_name,
                    item.get("frame", ""),
                    item.get("delay", ""),
                    item.get("unit", ""),
                ]
            )
    df_schedule = pd.DataFrame(schedule_rows, columns=schedule_columns)

    with pd.ExcelWriter(output_path) as writer:
        df_info.to_excel(writer, sheet_name="Info", index=False)
        df_matrix.to_excel(writer, sheet_name="Matrix", index=False)
        df_schedule.to_excel(writer, sheet_name="LIN Schedule", index=False)


def main():
    data = read_file_ldf("ATOM_LIN_Matrix_BCM-ALM_V4.0.0-20250121.ldf")

    print(extract_info(data=data))

    print(extract_nodes(data=data))

    pprint.pprint(extract_signals(data=data))

    pprint.pprint(extract_frames(data=data))

    pprint.pprint(extract_node_attributes(data=data))

    pprint.pprint(extract_schedule_tables(data=data))

    pprint.pprint(extract_signal_encoding_types(data=data))

    ldf_dicts_to_xlsx(
        extract_info(data=data),
        extract_nodes(data=data),
        extract_signals(data=data),
        extract_frames(data=data),
        extract_node_attributes(data=data),
        extract_schedule_tables(data=data),
        extract_signal_encoding_types(data=data),
    )


if __name__ == "__main__":
    main()
