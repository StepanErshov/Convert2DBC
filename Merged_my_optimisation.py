import pandas as pd
import streamlit as st
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Protection, Alignment
from openpyxl.utils import get_column_letter
from copy import copy
import os
from pathlib import Path
from io import BytesIO
from datetime import datetime
# Формирование папок
import create_directory
# Замер времени выполнения блоков
import time

def set_page_title():
    st.title("Release Convertor")

def get_uploaded_file():
    uploaded_file = st.file_uploader("Upload Domain Excel Matrix", type=["xlsx"])
    if uploaded_file:
        return uploaded_file

def process_matrix_sheet(df, domain_matrix, ecu_matrices, ecu_col_index):
    process_matrix_sheet_time_start = time.time()
    
    ecu_rows_to_copy = {}
    for ecu_name, ecu_index in ecu_col_index.items():
        rows_to_copy = df[df[ecu_name].notna()].index.tolist()
        ecu_rows_to_copy[ecu_name] = rows_to_copy
    
    print("Obtained ecu_rows_to_copy")
    
    domain_ws = domain_matrix['Matrix']

    for ecu_name, row_list_to_copy in ecu_rows_to_copy.items():
        row_to_paste = 1
        for row_idx in row_list_to_copy:
            for col_idx in range(1, domain_ws.max_column + 1):
                domain_matrix_cell = domain_ws.cell(row=row_idx+1, column=col_idx)
                ecu_matrix_cell = ecu_matrices[ecu_name]['Matrix'].cell(row=row_to_paste, column=col_idx)
                
                # Copy value
                ecu_matrix_cell.value = domain_matrix_cell.value
                
                # Copy formatting if it exists
                if domain_matrix_cell.has_style:
                    ecu_matrix_cell.font = copy(domain_matrix_cell.font)
                    ecu_matrix_cell.border = copy(domain_matrix_cell.border)
                    ecu_matrix_cell.fill = copy(domain_matrix_cell.fill)
                    ecu_matrix_cell.number_format = copy(domain_matrix_cell.number_format)
                    ecu_matrix_cell.protection = copy(domain_matrix_cell.protection)
                    ecu_matrix_cell.alignment = copy(domain_matrix_cell.alignment)
            row_to_paste +=1
        print(f'{ecu_name} processed')
        # ecu_matrices[ecu_name].save(f'output_ecus/{ecu_name}.xlsx')
        # print("saved")

    # Группировать строки по сообщениям
    row_idx = 2
    while row_idx <= new_matrix_ws.max_row:
        a_val = new_matrix_ws.cell(row=row_idx, column=1).value
        new_matrix_ws.row_dimensions[row_idx].height = 20
        if a_val:
            start_row = row_idx + 1
            end_row = start_row
            while end_row <= new_matrix_ws.max_row:
                i_val = new_matrix_ws.cell(row=end_row, column=9).value
                if i_val:
                    end_row += 1
                else:
                    break
            if end_row > start_row:
                for r in range(start_row, end_row):
                    new_matrix_ws.row_dimensions[r].outlineLevel = 1
                    new_matrix_ws.row_dimensions[r].hidden = True
                new_matrix_ws.row_dimensions[row_idx].collapsed = True
            row_idx = end_row
        else:
            row_idx += 1

    process_matrix_sheet_time_end = time.time() - process_matrix_sheet_time_start
    print("process_matrix_sheet_time_end", process_matrix_sheet_time_end)

    return ecu_matrices

def process_history_sheet(df, domain_matrix, ecu_matrices, ecu_col_index):

    process_history_sheet_time_start = time.time()

    if "History" not in domain_matrix.sheetnames:
        return False

    history_rows_to_copy = {}
    for ecu_name, ecu_index in ecu_col_index.items():
        rows_to_copy = df[df[ecu_name].notna()].index.tolist()
        ecu_rows_to_copy[ecu_name] = rows_to_copy

    ecu_rows_to_copy = {}
    for ecu_name, ecu_index in ecu_col_index.items():
        rows_to_copy = df[df[ecu_name].notna()].index.tolist()
        ecu_rows_to_copy[ecu_name] = rows_to_copy
    # Удалить строки, не соответствующие выбанному ecu
    ecu_history_ws = wb['History']
    history_col_idx = 6
    ecu_name = next(iter(ecu_col_index))
    for row_idx in range(ecu_history_ws.max_row, 1, -1):
        cell_value = ecu_history_ws.cell(row=row_idx, column=history_col_idx).value
        if cell_value is None or ecu_name not in str(cell_value):
            ecu_history_ws.delete_rows(row_idx)
    
    process_history_sheet_time_end = time.time() - process_history_sheet_time_start
    print("process_history_sheet_time_end", process_history_sheet_time_end)

    return ecu_wb

def identify_ecus(df):
    # Найти все ecu в предоставленной доменной матрице
    ecus = [
        col
        for col in df.columns
        if any(val in ["S", "R"] for val in df[col].dropna().unique())
        and col != "Unit\n单位"
    ]

    if not ecus:
        st.error(
            "🕭 No ECUs found in the file. Please check that the file contains columns with 'S' or 'R' values."
        )
        st.stop()

    return ecus

def get_ecu_version(wb, ecu_col_index):

    get_ecu_version_time_start = time.time()
    
    ecu_history_ws = wb['History']
    history_col_idx = 6
    ecu_name = next(iter(ecu_col_index))
    # Найти версию по первому снизу вхождению выбранного ecu
    for row_idx in range(ecu_history_ws.max_row, 1, -1):
        cell_value = ecu_history_ws.cell(row=row_idx, column=history_col_idx).value
        if ecu_name in str(cell_value):
            ecu_version = ecu_history_ws.cell(row=row_idx, column=1).value
            break
    
    get_ecu_version_time_end = time.time() - get_ecu_version_time_start
    print("get_ecu_version_time_end", get_ecu_version_time_end)

    return ecu_version

def get_domain_folder_name(ecu_base, domain_short):
    if ecu_base in ("SGW", "CGW", "ADCU"):
        domain_short = "SGW-CGW"
    domain_folder_name = {
        "BD" :      "04.01.01.Body CAN",
        "DG" :      "",
        "CH" :      "04.01.05.Chassis CANFD",
        "PT" :      "04.01.02.Powertrain CAN",
        "ET" :      "04.01.04.Entertainment CANFD",
        "DZ" :      "04.01.06.Demilitary zone CANFD",
        "SGW-CGW" : "04.01.07.CGW,SGW,ADCU"
    }


    return domain_folder_name[domain_short]

def get_ecu_folder_name(domain_folder_name, ecu_base):
    ecu_folder_name = ""
    for sub_folder_name in create_directory.creator.HIERARCHI[domain_folder_name]:
        print(ecu_base, sub_folder_name)
        if ecu_base in sub_folder_name:
            ecu_folder_name = sub_folder_name
            break
    if ecu_folder_name:
        return ecu_folder_name
    else:
        print("Sub folder for ECU doesn't match ECU name.", )
        # st.error("Sub folder for ECU doesn't match ECU name.")

def get_ecu_matrix_template(uploaded_file, ecu_col_index):
    col_index_to_letter = {i: get_column_letter(i) for i in range(1, 51)}
    ecu_wb = load_workbook(uploaded_file)
    matrix_ws_with_data = ecu_wb["Matrix"]
    history_ws_with_data = ecu_wb["History"]
    ecu_wb.remove(matrix_ws_with_data)
    ecu_wb.remove(history_ws_with_data)
    ecu_wb.create_sheet("Matrix")
    ecu_wb.create_sheet("History")
    
    # Сохранить ширину столбцов
    for col_idx in col_index_to_letter:
        if col_idx in ecu_col_index.values():
            ecu_wb["Matrix"].column_dimensions[col_index_to_letter[col_idx+1]].width = 2
        else:
            col_dimension = matrix_ws_with_data.column_dimensions[col_index_to_letter[col_idx]].width
            ecu_wb["Matrix"].column_dimensions[col_index_to_letter[col_idx]].width = col_dimension
    
    return ecu_wb

if __name__ == "__main__":
    set_page_title()
    uploaded_file = get_uploaded_file()
    # Доменный excel получен, дальше блок обработки для получения данных и распределение по папкам 
    if uploaded_file:
        try:
            print("File uploaded")
            st.info("File uploaded. Processing...")
            df = pd.read_excel(uploaded_file, sheet_name="Matrix")
            ecus = identify_ecus(df)
            progress_bar = st.progress(0)
            domain_matrix = load_workbook(uploaded_file)
            # Получить номер столбца для каждого ecu
            ecu_col_index = {ecu: df.columns.get_loc(ecu) for ecu in ecus}

            total_start = time.time()
            print("Ecu matrices start creating")
            # Создать пустые excel матрицы для каждого ecu
            ecu_matrices = {}
            for ecu_name in ecu_col_index:
                ecu_wb = get_ecu_matrix_template(uploaded_file, ecu_col_index)
                ecu_matrices[ecu_name] = ecu_wb
            print("Ecu matrices created")

            # Настроить лист "Matrix" под ecu
            ecu_matrices = process_matrix_sheet(df, domain_matrix, ecu_matrices, ecu_col_index)
            ecu_matrices = process_history_sheet(df, domain_matrix, ecu_matrices, ecu_col_index)
            if not ecu_wb:
                st.warning(f"🕱 'History' sheet not found in ECU {ecu}.")
            ecu_version = get_ecu_version(ecu_matrices, ecu_col_index)
            if not create_directory:
                st.warning("'create_directory' module is missing. Skipping file save.")
            domain_short = uploaded_file.name.split('_')[3]
            for ecu_name in ecu_col_index.keys():
                # Получить основу имени ecu
                if '_' in ecu_name:
                    ecu_base = ecu_name.split("_")[0]
                else:
                    ecu_base = ecu_name
                domain_folder_name = get_domain_folder_name(ecu_base, domain_short)
                ecu_folder_name = get_ecu_folder_name(domain_folder_name, ecu_base)
                if not ecu_folder_name:
                    st.warning(f"ECU folder name not found for {ecu_base}. Skipping.")
                # Сохранить ecu матрицу в папку
                find_path_time_start = time.time()
                date_str = datetime.now().strftime("%d%m%Y")
                output_ecu_filename = f"ATOM_CAN_MATRIX_{ecu_version}_{date_str}_{ecu}.xlsx"
                ecu_matrix_output_path = f"{create_directory.creator.PATH_DOC}\\{domain_folder_name}\\{ecu_folder_name}\\{output_ecu_filename}"
                find_path_time_end = time.time() - find_path_time_start
                print("find_path_time_end", find_path_time_end)

                ecu_save_time_start = time.time()
                ecu_wb.save(ecu_matrix_output_path)
                ecu_save_time_end = time.time() - ecu_save_time_start
                print("ecu_save_time_end", ecu_save_time_end)

            st.info(f"ECU {ecu}: processed in {ecu_proccesed_time:.2f} seconds")
            progress_bar.progress((idx + 1) / total_ecus)
            
            ecu_proccesed_time = time.time() - total_start
            print("ecu_proccesed_time", ecu_proccesed_time)

            st.success(
                f"Domain matrix ecu split completed, obtained {len(ecus)} ECUs."
            )
        except Exception as e:
            st.error(f"🕭 Error processing file: {e}")