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
