# def calculate_start_bit(start_byte, start_bit, byte_order, length):
#     """Calculate start bit considering byte order and signal length."""
#     if byte_order == "motorola":
#         return (start_byte * 8) + (7 - (start_bit % 8))
#     return (start_byte * 8) + start_bit

# def calculate_message_length(signals):
#     """Calculate minimum message length needed for all signals."""
#     max_bit = 0
#     for signal in signals:
#         end_bit = signal.start + signal.length
#         max_bit = max(max_bit, end_bit)
#     return (max_bit + 7) // 8


# start_bit = calculate_start_bit(
#     int(row["Start Byte"]),
#     int(row["Start Bit"]),
#     byte_order,
#     int(row["Length"])
# )


# used_bits = set()
# overlap_found = False
# for signal in signals:
#     start = signal.start
#     end = start + signal.length
#     for bit in range(start, end):
#         if bit in used_bits:
#             print(f"Warning: Overlapping bits in message {msg_name} (0x{frame_id:X}), signal {signal.name} (Bit {bit})")
#             overlap_found = True
#             break
#         used_bits.add(bit)
#     if overlap_found:
#         break

# if overlap_found:
#     print(f"Skipping message {msg_name} (0x{frame_id:X}) due to bit overlap")
#     continue

# message_length = calculate_message_length(signals)
# import os
# def get_file_info(file_name: str):
#         file_start = 'ATOM_CAN_Matrix_'
#         file_start1 = 'ATOM_CANFD_Matrix_' 
#         file_name_only = os.path.splitext(os.path.basename(file_name))[0]
#         if file_name_only.startswith(file_start1):
#             protocol = 'CANFD'
#             start_index = 0
#             parts = file_name_only[len(file_start1):].split('_')
#         elif file_name_only.startswith(file_start):
#             protocol = 'CAN'
#             start_index = 0
#             parts = file_name_only[len(file_start):].split('_')
#         else:
#             protocol = ''
#         if not (file_name_only.startswith(file_start) or file_name_only.startswith(file_start1)):
#             return None
#         start_index = file_name_only.find(file_start1)
#         if start_index != -1:
#             parts = file_name_only[start_index + len(file_start1):].split('_')
#         else:
#             parts = file_name_only[len(file_start):].split('_')
#         domain_name = parts.pop(0)
#         version_string = parts.pop(0)
#         if version_string.startswith('V'):
#             version = version_string[1:]
#             versions = version.split('.')
#             if len(versions) != 3:
#                 return None
#         else:
#             version = ''
#         file_date = parts.pop(0)
#         if len(parts) > 0:
#             if parts[0] == 'internal': # skip it
#                 parts.pop(0)
#             device_name = '_'.join(parts)
#         else:
#             device_name = ''

#         return {'version': version, 'date': file_date, 'device_name': device_name, 'domain_name': domain_name, "protocol": protocol}


# print(get_file_info("ATOM_CANFD_Matrix_CH_internal_V2.2.0_ACU"))