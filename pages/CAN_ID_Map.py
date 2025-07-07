import streamlit as st
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Font
from datetime import datetime
from io import BytesIO
import re
import cantools

def set_page_config():
    st.title("🔥CAN ID Map")

def files_upload():
    uploaded_files = st.file_uploader("Load domain matrices or DBC", type=["xlsx", "dbc"], accept_multiple_files=True)
    uploaded_files_by_format = {}
    excel_files = []
    dbc_files = []
    if uploaded_files:
        st.success(f"Uploaded matrices: {len(uploaded_files)}")
        for file in uploaded_files:
            if file.name.endswith('.xlsx'):
                excel_files.append(file)
            else:
                dbc_files.append(file)
            st.write(file.name)
        uploaded_files_by_format["xlsx"] = excel_files
        uploaded_files_by_format["dbc"] = dbc_files

        return uploaded_files_by_format
    return 0

def get_format_splitted_files(uploaded_files):
    if uploaded_files:
        excel_files = uploaded_files["xlsx"]
        dbc_files = uploaded_files["dbc"]

        return excel_files, dbc_files
    return 0, 0

def input_version():
    version = st.text_input("Enter version (e.g., V3.1.1):")
    pattern = r"^V\d\.\d\.\d$"
    if version:
        if not re.match(pattern, version):
            st.error("Invalid format! Please enter as V1.2.3")
    else:
        version = "VNone"
    return version

def get_excel_2_df(uploaded_files):
    if uploaded_files:
        pd_df_matrices = []
        # Получить датафреймы для каждого файла
        for file in uploaded_files:
            df = pd.read_excel(file, sheet_name="Matrix")
            necessary_columns = [0, 2, 4]
            df = df.iloc[:, necessary_columns]
            message_name_column = df.columns[0]
            df.dropna(subset=[message_name_column], inplace=True)
            df.columns = ['message name', 'message id', 'message cycle time']
            df['message id'] = df['message id'].str[2:]
            pd_df_matrices.append(df)

        return pd_df_matrices
    return 0

def get_dbc_2_df(uploaded_files):
    if uploaded_files:
        pd_df_matrices = []
        # Получить датафреймы для каждого файла
        for file in uploaded_files:
            dbc_content = file.read().decode('utf-8')
            db = cantools.database.load_string(dbc_content, 'dbc')
            data = []
            for message in db.messages:
                message_id = hex(message.frame_id).upper()[2:]
                if len(message_id) == 2:
                    message_id = '0' + message_id
                data.append({
                    'message name': message.name,
                    'message id': message_id,
                    'message cycle time': message.cycle_time
                })
            df = pd.DataFrame(data)
            pd_df_matrices.append(df)

        return pd_df_matrices
    return 0

def get_merged_df(df_excel, df_dbc):
    if not isinstance(df_excel, int) and not isinstance(df_dbc, int):
        excel_dbc_df = df_excel + df_dbc
        merged_df = pd.concat(excel_dbc_df).drop_duplicates().reset_index(drop=True)
    elif not isinstance(df_excel, int):
        merged_df = pd.concat(df_excel).drop_duplicates().reset_index(drop=True)
    elif not isinstance(df_dbc, int):
        merged_df = pd.concat(df_dbc).drop_duplicates().reset_index(drop=True)
    else:
        return 0
    return merged_df

def stylise_cell(cell, msg_cycle_time):
    cycle_time_to_fill = {
        10.0:   PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid"),  # Red
        20.0:   PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid"),  # Orange/Gold
        50.0:   PatternFill(start_color="7030A0", end_color="7030A0", fill_type="solid"),  # Purple
        100.0:  PatternFill(start_color="33cc33", end_color="33cc33", fill_type="solid"),  # Green
        200.0:  PatternFill(start_color="66ff66", end_color="66ff66", fill_type="solid"),  # Light Green
        ">200.0": PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid"),  # Blue
        "nan":  PatternFill(start_color="FF66CC", end_color="FF66CC", fill_type="solid"),  # Pink
    }

    custom_font = Font(
        name='Arial',
        size=12,
    )

    cell.font = custom_font
    if float(msg_cycle_time) > 200:
        cell.fill = cycle_time_to_fill[">200.0"]
    elif pd.isna(msg_cycle_time):
        cell.fill = cycle_time_to_fill["nan"]
    else:
        cell.fill = cycle_time_to_fill[msg_cycle_time]

def get_check_result_ws(df, CAN_ID_Map):
    if CAN_ID_Map and not isinstance(df, int):
        id_map_ws = CAN_ID_Map["ATOM_ID Map"]
        check_result_ws = CAN_ID_Map.copy_worksheet(id_map_ws)
        check_result_ws.title = "CheckResult"

        return check_result_ws
    return 0

def get_overlays_df(df):
    if not isinstance(df, int):
        overlays_df = df[df['message id'].duplicated(keep=False)]

        return overlays_df
    return 0

def show_overlays(overlays_df):
    if not isinstance(overlays_df, int):
        if not overlays_df.empty:
            st.error("Overlayed ids:")
            st.write(overlays_df)

def get_multi_id_messages(df):
    if not isinstance(df, int):
        multi_id_messages = {}
        multi_id_messages_df = df[df['message name'].duplicated(keep=False)]
        multi_id_messages = multi_id_messages_df.groupby('message name')['message id'].unique().apply(list).to_dict()

        return multi_id_messages
    return 0

def show_multi_id_messages(multi_id_messages):
    if multi_id_messages:
        st.error("Multi ID messages:")
        st.write(multi_id_messages)

def generate_CAN_ID_Map(template_path, df):
    if template_path and not isinstance(df, int):
        # Загрузить шаблон
        CAN_ID_Map = load_workbook(template_path)
        # Добавить лист CheckResult
        check_result_ws = get_check_result_ws(df, CAN_ID_Map)
        # Получить датафрейм с наложенными id
        overlays_df = get_overlays_df(df)
        # Вывести в интерфейс наложенные id
        show_overlays(overlays_df)
        # Получить сообщения с неоднозначными id
        multi_id_messages = get_multi_id_messages(df)
        # Вывести в интерфейс сообщения с неоднозначными id
        show_multi_id_messages(multi_id_messages)
        history_ws = CAN_ID_Map['History']
        id_map_ws = CAN_ID_Map["ATOM_ID Map"]
        # Подготовить словарь для конвертации msg_id в id столбца excel 
        hex_column_id = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9', 'A', 'B', 'C', 'D', 'E', 'F']
        excel_column_id =  ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q']
        hex_to_column = dict(zip(hex_column_id, excel_column_id))
        # Заполнить дату создания
        history_ws['B2'] = datetime.now().strftime("%d/%m/%Y")
        # Заполнение таблицы
        for index, row in df.iterrows():
            msg_name = row.iloc[0]
            msg_id = row.iloc[1]
            msg_id_row = int(row.iloc[1][:2], 16)
            msg_id_column = hex_to_column[row.iloc[1][2:]]
            msg_cycle_time = row.iloc[2]
            # Заполнение msg_name
            cell_id = f'{msg_id_column}{msg_id_row+1}'
            # Проверка на наложение и заполнение ячеек
            if msg_name in overlays_df.iloc[:, 0].values or msg_name in multi_id_messages:
                if not id_map_ws[cell_id].value:
                    id_map_ws[cell_id] = msg_name
                    check_result_ws[cell_id] = msg_name
                else:
                    id_map_ws[cell_id] = str(id_map_ws[cell_id].value) + f'\n{msg_name}'
                    check_result_ws[cell_id] = str(check_result_ws[cell_id].value) + f'\n{msg_name}'
            else:
                id_map_ws[cell_id] = msg_name
                stylise_cell(id_map_ws[cell_id], msg_cycle_time)

        CAN_ID_Map.save("CAN_ID_Map.xlsx")
        CAN_ID_Map.close()

        return CAN_ID_Map
    return 0

def download_CAN_ID_Map(CAN_ID_Map, version):
    if CAN_ID_Map:
        output_path = f"CANID Design_ATOM_{version}-{datetime.now().strftime("%Y%m%d")}.xlsx"
        buffer = BytesIO()
        CAN_ID_Map.save(buffer)
        buffer.seek(0)
        st.download_button(
            label="Download CAN ID map",
            data=buffer,
            file_name=output_path,
            mime="application/octet-stream",
            type="primary",
            help="Press to download .xlsx CAN ID map"
        )


def main():
    try:
        template_path = "pages/CAN_ID_Map_template.xlsx"

        # Озаглавить страницу
        set_page_config()
        # Предоставить форму для загрузки файлов
        uploaded_files = files_upload()
        # Получить файлы, разделенные на xlsx и dbc форматы
        excel_files, dbc_files = get_format_splitted_files(uploaded_files)
        # Предоставить поле для ввода версии таблицы
        version = input_version()
        # Перевести загруженные таблицы в датафрейм
        df_excel = get_excel_2_df(excel_files)
        # Перевести загруженные DBC в датафрейм
        df_dbc = get_dbc_2_df(dbc_files)
        # Совместить excel и dbc датафреймы
        merged_df = get_merged_df(df_excel, df_dbc)
        # Сгенерировать id таблицу
        CAN_ID_Map = generate_CAN_ID_Map(template_path, merged_df)
        # Скачать результат
        download_CAN_ID_Map(CAN_ID_Map, version)


    except Exception as e:
        st.error(f"Error occured: {str(e)}")
        st.stop()

if __name__ == "__main__":
    main()