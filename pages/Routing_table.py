import streamlit as st
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Border, Side, PatternFill, Font, Alignment
from copy import copy
from datetime import datetime
from io import BytesIO
import json
import itertools
 
def set_page_config():
    st.title("Routing Table")
 
def gateway_selection():
    return st.radio(
        "Choose gateway:",
        ("SGW", "CGW")
    )
 
def files_upload():
    uploaded_files = st.file_uploader("Load domain matrices", type=["xlsx"], accept_multiple_files=True)
    if uploaded_files:
        st.success(f"✅ Uploaded matrices: {len(uploaded_files)}")
        for file in uploaded_files:  # ограничиваем вывод
            st.write(file.name)
    return uploaded_files
 
def get_pd_data(uploaded_files):
    if uploaded_files:
        pd_df_matrices = {}
        # Для каждого файла получить пару: имя - датафрейм
        for file in uploaded_files:
            df = pd.read_excel(file, sheet_name="Matrix")
            message_id_column_name = df.columns[2]
            df.dropna(subset=[message_id_column_name], inplace=True)
            pd_df_matrices[file.name] = df
 
        return pd_df_matrices
    return 0
 
def get_routing_table_template_path(input_path, output_path, uploaded_files):
    if uploaded_files:
        # Загрузить сырой шаблон
        routing_table_template = load_workbook(input_path)
        with open('./pages/template_values.json', 'r', encoding='utf-8') as template_values_json:
            template_values = json.load(template_values_json)
        all_domains = {"BD", "DG", "PT", "CH", "DZ", "ET", "SGW"}
        routed_domains = []
        # Определить домены загруженных матриц
        for file in uploaded_files:
            for domain in all_domains:
                if domain in file.name:
                    routed_domains.append(domain)
         
        real_values = {
            "release date":         datetime.now().strftime("%Y.%m.%d"),
            "source domains":       routed_domains,
            "target domains":       routed_domains
        }
 
        sheet_names = routing_table_template.sheetnames
        worksheets = {sheet_names[i] : routing_table_template.worksheets[i] for i in range(len(sheet_names))}
        # Заполнить первый лист (обложка)
        for row in worksheets[sheet_names[0]].iter_rows():
            for cell in row:
                    if cell.value and isinstance(cell.value, str):
                        if "Current date" in cell.value:
                            cell.value = cell.value.replace(template_values["release date"], real_values["release date"])
        # Заполнить второй лист (история изменений)
        for row in worksheets[sheet_names[1]].iter_rows():
            for cell in row:
                    if cell.value and isinstance(cell.value, str):
                        key = next((k for k, v in template_values.items() if v == cell.value), None)
                        if key:
                            if isinstance(real_values[key], str):
                                cell.value = real_values[key]
                            else:
                                cell.value = ', '.join(real_values[key])
 
        routing_table_template.save(output_path)
        routing_table_template.close()
         
        return output_path
    return 0
 
def calculate_routing_table_data(pd_df_matrices, gateway):
    if pd_df_matrices and gateway:
        # Для хранения всех message_id для каждой матарицы
        matrix_message_ids = {}
        # Для хранения общих между парой матриц фреймов
        source_target_common_ids = {}
        # Для хранения маршрутизируемых сообщений для каждой пары имен матриц
        routing_table_data = {}
        # Получение id всех сообщений рассматриваемой матрицы
        for matrix_name, pd_df_matrix in pd_df_matrices.items():
            matrix_message_ids[matrix_name] = pd_df_matrix.iloc[:, 2]
        # Получение всех пар матриц для построения таблицы маршрутов
        source_target_data = list(itertools.permutations(pd_df_matrices.keys(), 2))
        for source_taget in source_target_data:
            source = source_taget[0]
            target = source_taget[1]
            # Получение общих сообщений для рассматриваемой пары матриц
            routed_ids = pd.merge(matrix_message_ids[source], matrix_message_ids[target], how='inner')
            # Для каждой пары маттриц определяются id маршрутизируемых сообщений
            source_target_common_ids[source_taget] = routed_ids
            # Определяются сообщения маршрутизируемые из матрицы источника
            source_matrix = pd_df_matrices[source]
            message_id_column_name = source_matrix.columns[2]
            source_gateway_column = source_matrix.filter(like=gateway).columns
            # Проверка на наличие выбранного шлюза в ECU матрицы источника
            if len(source_gateway_column) == 0:
                st.error(f"'{gateway}' not in '{source}'. Check gateway.")
                st.stop()
            routed_ids_list = routed_ids[message_id_column_name].tolist()
            source_matrix_routed_messages = source_matrix[(source_matrix[message_id_column_name].isin(routed_ids_list)) & (source_matrix[source_gateway_column[0]] == 'R')]
            # Определяются сообщения маршрутизируемые в целевую матрицу
            target_matrix = pd_df_matrices[target]
            message_id_column_name = target_matrix.columns[2]
            target_gateway_column = target_matrix.filter(like=gateway).columns
            # Проверка на наличие выбранного шлюза в ECU матрицы получателя
            if len(target_gateway_column) == 0:
                st.error(f"'{gateway}' not in '{target}'. Check gateway.")
                st.stop()
            target_matrix_routed_messages = target_matrix[(target_matrix[message_id_column_name].isin(routed_ids_list)) & (target_matrix[target_gateway_column[0]] == 'S')]
            # Определяются маршрутизируемые сообщения между данными матрицами
            matrices_routed_messages = pd.merge(source_matrix_routed_messages, target_matrix_routed_messages, how='inner')
            matrices_routed_messages = matrices_routed_messages[[matrices_routed_messages.columns[0], matrices_routed_messages.columns[2]]]
            # Запись в словарь маршрутизируемых между данными матрицами сообщений по ключу рассматриваемой пары матриц
            routing_table_data[source_taget] = matrices_routed_messages
             
        return routing_table_data
    return 0
 
def generate_routing_table(routing_table_data, routing_table_template_path, gateway):
    if (routing_table_data) and (routing_table_template_path) and (gateway):
        generate_btn = st.button("Generate")
        if generate_btn:
            # Загрузить подготовленный шаблон
            routing_table = load_workbook(routing_table_template_path)
            table_headers = ['Signal Name', 'Message Name', 'Message ID', 'Signal Name', 'Message Name', 'Message ID', 'Routing Type', 'Gateway ECU', 'Change Record']
            start_row = 3
            start_col = 1
            merge_count = 3
            route_table_worksheet = routing_table.worksheets[-1]
            current_col = start_col
            current_raw = start_row
            # начало стили (нужно перенести в изолированное место)
            thin_border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
            blue_fill = PatternFill(start_color='00ccff', fill_type='solid')
            orange_fill = PatternFill(start_color='ff9900', fill_type='solid')
            yellow_fill = PatternFill(start_color='ffff99', fill_type='solid')
            green_fill = PatternFill(start_color='ccffcc', fill_type='solid')
 
            text_center_alignment = Alignment(horizontal='center', vertical='center')
 
            custom_font_bold = Font(
                name='Calibri',
                size=12,
                bold=True,
            )
 
            custom_font = Font(
                name='Calibri',
                size=12,
            )
            # конец стили (нужно перенести в изолированное место)
            for matrix_pair, route_data in routing_table_data.items():
                # Определить данные строки с источником и получателем
                source_taget_header = [f'Source: {matrix_pair[0]}',
                                        f'Target: {matrix_pair[1]}',
                                        None]
                # Заполнить строку с источником и получателем
                for i in range(merge_count):
                    cell = route_table_worksheet.cell(row=current_raw, column=current_col, value=source_taget_header[i])
                    cell.border = thin_border
                    cell.font = custom_font_bold
                    cell.alignment = text_center_alignment
                    route_table_worksheet.merge_cells(
                        start_row=current_raw,
                        start_column=current_col,
                        end_row=current_raw,
                        end_column=current_col + merge_count - 1
                    )
                    current_col += merge_count
 
                current_raw += 1
                current_col = start_col
                # Заполнить заголовки таблицы маршрутизации для рассматриваемой пары
                for table_header in table_headers:
                    cell = route_table_worksheet.cell(row=current_raw, column=current_col, value=table_header)
                    cell.border = thin_border
                    cell.font = custom_font
                    cell.fill = blue_fill
                    cell.alignment = text_center_alignment
                    current_col += 1
 
                current_raw += 1
                current_col = start_col
                # Перевести датафрейм в строки
                for route_table_row_data in dataframe_to_rows(route_data, index=False, header=False):
                    # Заполнение полей Message Name и Message ID отправителя и получателя
                    for cell_value in route_table_row_data:
                        # Заполнение полей Message Name и Message ID отправителя
                        cell = route_table_worksheet.cell(row=current_raw, column=current_col+1, value=cell_value)
                        cell.border = thin_border
                        cell.font = custom_font
                        cell.fill = green_fill
                        cell.alignment = text_center_alignment
                        # Заполнение полей Message Name и Message ID получателя
                        cell = route_table_worksheet.cell(row=current_raw, column=current_col+4, value=cell_value)
                        cell.border = thin_border
                        cell.font = custom_font
                        cell.fill = green_fill
                        cell.alignment = text_center_alignment
                        current_col += 1
                    current_col = start_col
                    # Заполнение поля Signal Name отправителя
                    cell = route_table_worksheet.cell(row=current_raw, column=current_col, value=None)
                    cell.border = thin_border
                    cell.font = custom_font
                    cell.fill = orange_fill
                    cell.alignment = text_center_alignment
                    # Заполнение поля Signal Name получателя
                    cell = route_table_worksheet.cell(row=current_raw, column=current_col+3, value=None)
                    cell.border = thin_border
                    cell.font = custom_font
                    cell.fill = orange_fill
                    cell.alignment = text_center_alignment
                    # Заполнение поля Routing Type
                    cell = route_table_worksheet.cell(row=current_raw, column=current_col+6, value="Message")
                    cell.border = thin_border
                    cell.font = custom_font
                    cell.fill = yellow_fill
                    cell.alignment = text_center_alignment
                    # Заполнение поля Gateway ECU
                    cell = route_table_worksheet.cell(row=current_raw, column=current_col+7, value=gateway)
                    cell.border = thin_border
                    cell.font = custom_font
                    cell.fill = yellow_fill
                    cell.alignment = text_center_alignment
                    # Заполнение поля Change Record
                    cell = route_table_worksheet.cell(row=current_raw, column=current_col+8, value=None)
                    cell.border = thin_border
                    cell.font = custom_font
                    cell.fill = yellow_fill
                    cell.alignment = text_center_alignment
                    # переход на следующую строку
                    current_col = start_col
                    current_raw += 1
                 
                current_raw += 1
             
            # routing_table.save("routing_table.xlsx")
            routing_table.close()
        else:
            routing_table = 0
 
        return routing_table
    return 0
 
def download_routing_table(routing_table, gateway):
    if routing_table:
        output_path = f"RoutingMAP-{gateway}_VNone_{datetime.now().strftime("%Y%m%d")}.xlsx"
        buffer = BytesIO()
        routing_table.save(buffer)
        buffer.seek(0)
        st.download_button(
            label="Download routing map",
            data=buffer,
            file_name=output_path,
            mime="application/octet-stream",
            type="primary",
            help="Press to download .xlsx RoutingMap"
        )
 
def main():
    try:
        # Озаглавить страницу
        set_page_config()
        # Предоставить форму для загрузки файлов
        uploaded_files = files_upload()
        # Предоставить выбор шлюза
        gateway = gateway_selection()
        # Подготовить шаблон и получить его директорию
        routing_table_template_path = get_routing_table_template_path("./pages/routing_table_template.xlsx", f"routing_table_template{datetime.now().strftime("%Y_%m_%d")}.xlsx", uploaded_files)
        # Считать загруженные файлы в pandas датафреймы
        pd_df_matrices = get_pd_data(uploaded_files)
        # Обработать загруженные данные для получения данных для заполнения таблицы маршрутизации
        routing_table_data = calculate_routing_table_data(pd_df_matrices, gateway)
        # Заполнить таблицу маршрутизации с требуемым форматированием
        routing_table = generate_routing_table(routing_table_data, routing_table_template_path, gateway)
        # Предоставить конфигурацию для загрузки выходного файла
        download_routing_table(routing_table, gateway)
 
    except Exception as e:
        st.error(f"Error occured: {str(e)}")
        st.stop()
 
if __name__ == "__main__":
    main()