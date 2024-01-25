
from time import perf_counter
from random import randint
from zipfile import BadZipFile
from re import findall, IGNORECASE
from io import BytesIO

import streamlit as st
from streamlit import session_state
from docx import Document
from docx.opc.exceptions import PackageNotFoundError
from pandas import DataFrame

from vezamodule import SUPPORTED_EXCTENTIONS_FOR_BLANK, ideal_message, Blank
from HOVS import hovs, two_excel_calc, sort_rule

# streamlit run c:\Users\ovchinnikov\Documents\Python\30_HOVS\HOVS_WEB.py
# streamlit run c:\users\ovchinnikov\documents\python\30_hovs\hovs_web.py
# streamlit run c:\users\krinitsin.da\desktop\stprogs\hovs_web.py
# streamlit run e:\python\30_hovs\hovs_web.py

def main_part():
    """_summary_
    """

    if 'run' not in session_state:
        session_state['run'] = 0
    if 'all_files' not in session_state:
        session_state['all_files'] = list()
    if 'result_df' not in session_state:
        session_state['result_df'] = DataFrame()
    if 'upload_files' not in session_state:
        session_state['upload_files'] = 'start' + str(randint(1, 1000) * randint(1, 9))

    VERSION_HOVS = '0.2.1.'

    tab_HOVS, tab_gaba = st.tabs(('ХОВС', 'Габариты'))

    with tab_HOVS:
        st.title(f'ХОВС. Версия {VERSION_HOVS}')

        all_files_hovs = sorted(st.file_uploader('Бланки для ХОВСа', SUPPORTED_EXCTENTIONS_FOR_BLANK[:2], True, session_state['upload_files'], 'Просто перемещайте конкретно в это меню все файлы, которые нужно обработать, дальше программа всё сделает, по завершении надо будет нажать на кнопку "Выгрузить ХОВС"'), key=lambda x: x.name)
        result = []

        if st.button('Очистить список файлов'):
            session_state['upload_files'] = 'start' + str(randint(1, 1000) * randint(1, 9))
            if st.button('Подтвердить очистку'):
                pass

        # findall_short = lambda mask:findall(mask, info, IGNORECASE)[0]

        if all_files_hovs:
            count_hovs = 0
            progress_bar_info_hovs = st.progress(1.0, 'Загрузите файл(ы)')
            progress_bar_hovs = st.progress(count_hovs, 'Загрузите файл(ы)')
            if session_state['all_files'] != all_files_hovs:
                session_state['all_files'] = all_files_hovs
                session_state['result_df'] = DataFrame()
                session_state['run'] = 0

            if session_state['run'] == 0:
                time_start = perf_counter()
                all_files_count_hovs = len(all_files_hovs)
                for file in all_files_hovs:
                    count_hovs += 1
                    progress_bar_info_hovs.progress(1.0, f"Обрабатывается файл {file.name}")
                    try:
                        docx_file = Document(file)
                    except (PackageNotFoundError, ValueError, BadZipFile):
                        st.write(Exception(f"Сохраните файл {file.name} в docx-формате и попробуйте ещё раз!"))
                    else:
                        result += hovs(Blank(docx_file))

                    progress_bar_hovs.progress(count_hovs / all_files_count_hovs, ideal_message(count_hovs, all_files_count_hovs, 'файлов', time_start, True, True))
                    pass
                
                session_state['result_df'] = DataFrame(sort_rule(result), columns=result[0].keys())
                session_state['run'] += 1
            st.download_button('Выгрузить ХОВС', two_excel_calc(session_state['result_df']).getvalue(), 'ХОВС.xlsx', 'xlsx')

    with tab_gaba:
        result_gaba_list = []

        all_files_gaba = st.file_uploader('Кидайте бланки сюда', SUPPORTED_EXCTENTIONS_FOR_BLANK[:2], True, 'all_files_gaba')
        if all_files_gaba:
            count_gaba = 0
            progress_bar_info_gaba = st.progress(1.0, 'Загрузите файл(ы)')
            progress_bar_gaba = st.progress(count_gaba, 'Загрузите файл(ы)')
            time_start = perf_counter()
            all_files_count_gaba = len(all_files_gaba)

            for file in all_files_gaba:
                count_gaba += 1
                progress_bar_info_gaba.progress(1.0, f"Обрабатывается файл {file.name}")
                try:
                    main_blank = Blank(Document(file))
                except (PackageNotFoundError, ValueError, BadZipFile) as err:
                    st.write(Exception(f"Сохраните файл {file.name} в docx-формате и попробуйте ещё раз!"))
                except Exception as err:
                    st.write(err)
                else:
                    # st.write(main_blank)
                    curr_itogo = []
                    for header, info in main_blank.ALL_MAIN_INFO.items():
                        res = findall(r'bфр=(.+?)мм; hфр=(.+?)мм; L=(.+?)мм; M=(.+?)кг;', info + ';', IGNORECASE)
                        # st.write(res)
                        if res:
                            res = res[0]
                            curr_itogo.extend(map(int, (res[1], res[0], res[2], res[3])))
                    result_gaba_list.append([main_blank.main_information['Название'], len(curr_itogo) / 4] + curr_itogo)
                    # st.write(curr_itogo, itogo)
                    pass
                progress_bar_gaba.progress(count_gaba / all_files_count_gaba, ideal_message(count_gaba, all_files_count_gaba, 'файлов', time_start, True, True))
            with st.expander('Проверить таблицу'):
                st.table(result_gaba_list)
            st.write(len(max(result_gaba_list, key=len)))
            res_bio = BytesIO()
            DataFrame(result_gaba_list, columns=['Название системы', 'Кол-во блоков'] + ['hфр, мм', 'bфр, мм', 'L, мм', 'M, кг'] * ((max(len(s) for s in result_gaba_list) - 2) // 4)).to_excel(res_bio, index=False)
            st.download_button('Загрузите', res_bio.getvalue(), 'Результаты.xlsx', 'xlsx')
    pass

if __name__ == "__main__":
    st.set_page_config(layout="wide")
    main_part()