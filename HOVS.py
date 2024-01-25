from re import findall, IGNORECASE
from io import BytesIO
from itertools import takewhile
from copy import deepcopy
from pathlib import Path

from pandas import DataFrame, ExcelWriter
import xlsxwriter

from vezamodule import print_debug_mode_on as pdmo, Blank

# streamlit run c:\Users\ovchinnikov\Documents\Python\30_HOVS\HOVS.py
# streamlit run c:\users\ovchinnikov\documents\python\30_hovs\hovs.py
# streamlit run c:\users\krinitsin.da\desktop\stprogs\hovs.py
# streamlit run e:\python\30_hovs\hovs.py

# key_words = {
#     "1.3" : "клапан" if not main_blank.IS_OBPROM else "стам",
#     "1.4" : "фильтр ",
#     "1.5" : "нагрев",
#     "1.6" : "воздухонагреватель",
#     "1.7" : "нагрев",
#     "1.8" : "воздухоохладитель канальный " if main_blank.IS_CHANAL else "воздухоохладитель жидкостный",
#     "1.9" : "воздухоохладитель непосредственного",
#     "1.10" : "теплоутилизатор",
#     "1.11" : "камера увлажнения",
#     "1.12" : "вентилятор" if not main_blank.IS_OBPROM and not main_blank.IS_INDUST else ALL_MAIN_INFO[0][0].lower(),
#     "1.13" : "вентилятор" if not main_blank.IS_OBPROM and not main_blank.IS_INDUST else ALL_MAIN_INFO[0][0].lower(),
# }

def two_excel_calc(df:DataFrame) -> BytesIO:
    """_summary_

    Args:
        df (DataFrame): _description_

    Returns:
        BytesIO: _description_
    """
    df = df.to_dict()
    
    try:
        df.pop('Unnamed: 0', None)
    except:
        pass

    for key in df.keys():
        df[key] = [None, None, None, None] + [df[key][i] for i in df[key].keys()]
    df = list(zip(*([i] + df[i] for i in df.keys())))
    df = [list(ite) for ite in df]

    lengths = list(map(lambda x: x / 2, (15, 10, 50, 30, 15, 20, 15, 15, 30, 10, 15,30, 10, 10, 10, 20, 15, 15, 30, 10, 15, 15, 10, 10, 20, 10, 15, 15, 30, 10, 15, 30, 10, 10, 10, 20, 15, 30, 10, 10, 20, 10, 15, 30)))
    output = BytesIO()
    writer =  ExcelWriter(output, engine='xlsxwriter')
    workbook = writer.book
    border_fmt = workbook.add_format({'bottom':1, 'top':1, 'left':1, 'right':1})
    DataFrame(df).to_excel(writer, index=False, header=False,sheet_name='Sheet1')
    worksheet = writer.sheets['Sheet1']
    worksheet.conditional_format(xlsxwriter.utility.xl_range(0, 0, len(df)-1, len(lengths)-1), {'type': 'no_errors', 'format': border_fmt})
    format1 = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'text_wrap': True, "font_name": 'GOST type A', "font_size":12, 'italic' : True})
    for i in range(len(df)):
        worksheet.set_row(i, 15)
    ki = 0
    for idx in range(44):
        max_len = lengths[ki]
        ki += 1
        worksheet.set_column(idx,idx,max_len, format1)
    m_ranges = [
        (f'A1:A5', df[0][0], format1),
        (f'B1:B5', df[0][1], format1),
        (f'C1:C5', df[0][2], format1),
        (f'D1:D5', df[0][3], format1),
        (f'E1:K1', df[0][4].split("_")[0], format1),
        (f'E2:E5', df[0][4].split("_")[1], format1),
        (f'F2:F5', df[0][5].split("_")[1], format1),
        (f'G2:G5', df[0][6].split("_")[1], format1),
        (f'H2:H5', df[0][7].split("_")[1], format1),
        (f'I2:K2', df[0][8].split("_")[1], format1),
        (f'I3:I5', df[0][8].split("_")[2], format1),
        (f'J3:J5', df[0][9].split("_")[2], format1),
        (f'K3:K5', df[0][10].split("_")[2], format1),
        (f'L1:R1', df[0][11].split("_")[0], format1),
        (f'L2:L5', df[0][11].split("_")[1], format1),
        (f'M2:M5', df[0][12].split("_")[1], format1),
        (f'N2:O3', df[0][13].split("_")[1], format1),
        (f'N4:N5', df[0][13].split("_")[2], format1),
        (f'O4:O5', df[0][14].split("_")[2], format1),
        (f'P2:P5', df[0][15].split("_")[1], format1),
        (f'Q2:R3', df[0][16].split("_")[1], format1),
        (f'Q4:Q5', df[0][16].split("_")[2], format1),
        (f'R4:R5', df[0][17].split("_")[2], format1),
        (f'S1:AB1', df[0][18].split("_")[0], format1),
        (f'S2:S5', df[0][18].split("_")[1], format1),
        (f'T2:T5', df[0][19].split("_")[1], format1),
        (f'U2:V3', df[0][20].split("_")[1], format1),
        (f'U4:U5', df[0][20].split("_")[2], format1),
        (f'V4:V5', df[0][21].split("_")[2], format1),
        (f'W2:X3', df[0][22].split("_")[1], format1),
        (f'W4:W5', df[0][22].split("_")[2], format1),
        (f'X4:X5', df[0][23].split("_")[2], format1),
        (f'Y2:Y5', df[0][24].split("_")[1], format1),
        (f'Z2:Z5', df[0][25].split("_")[1], format1),
        (f'AA2:AB3', df[0][26].split("_")[1], format1),
        (f'AA4:AA5', df[0][26].split("_")[2], format1),
        (f'AB4:AB5', df[0][27].split("_")[2], format1),
        (f'AC1:AE1', df[0][28].split("_")[0], format1),
        (f'AC2:AC5', df[0][28].split("_")[1], format1),
        (f'AD2:AD5', df[0][29].split("_")[1], format1),
        (f'AE2:AE5', df[0][30].split("_")[1], format1),
        (f'AF1:AK1', df[0][31].split("_")[0], format1),
        (f'AF2:AF5', df[0][31].split("_")[1], format1),
        (f'AG2:AG5', df[0][32].split("_")[1], format1),
        (f'AH2:AI3', df[0][33].split("_")[1], format1),
        (f'AH4:AH5', df[0][33].split("_")[2], format1),
        (f'AI4:AI5', df[0][34].split("_")[2], format1),
        (f'AJ2:AJ5', df[0][35].split("_")[1], format1),
        (f'AK2:AK5', df[0][36].split("_")[1], format1),
        (f'AL1:AQ1', df[0][37].split("_")[0], format1),
        (f'AL2:AL5', df[0][37].split("_")[1], format1),
        (f'AM2:AM5', df[0][38].split("_")[1], format1),
        (f'AN2:AN5', df[0][39].split("_")[1], format1),
        (f'AO2:AQ2', df[0][40].split("_")[1], format1),
        (f'AO3:AO5', df[0][40].split("_")[2], format1),
        (f'AP3:AP5', df[0][41].split("_")[2], format1),
        (f'AQ3:AQ5', df[0][42].split("_")[2], format1),
        (f'AR1:AR5', df[0][43], format1)
    ]
    for item in m_ranges:
        worksheet.merge_range(*item)

    # A, D, AR
    q_ranges = []
    range_beg, range_eng = 5, 5
    value_syst, valuer_name, value_qwer = df[range_beg][0], df[range_beg][3], df[range_beg][-1]
    for i in range(range_beg, len(df)):
        if df[i][0] == value_syst and df[i][-1] == value_qwer:
            range_eng = i
        else:
            q_ranges += [
                (f'A{range_beg+1}:A{range_eng+1}', value_syst, format1),
                (f'D{range_beg+1}:D{range_eng+1}', valuer_name, format1),
                (f'AR{range_beg+1}:AR{range_eng+1}', value_qwer, format1),
            ]
            value_syst, valuer_name, value_qwer, range_beg, range_eng = df[i][0], df[i][3], df[i][-1], i, i

    q_ranges += [
        (f'A{range_beg+1}:A{range_eng+1}', value_syst, format1),
        (f'D{range_beg+1}:D{range_eng+1}', valuer_name, format1),
        (f'AR{range_beg+1}:AR{range_eng+1}', value_qwer, format1),
    ]

    for item in q_ranges:
        worksheet.merge_range(*item)
    writer.close()
    return output

def hovs(main_blank:Blank):
    """_summary_

    Args:
        mask (str): _description_

    Returns:
        _type_: _description_
    """

    def findall_short(mask:str):
        """_summary_

        Args:
            mask (str): _description_

        Returns:
            _type_: _description_
        """
        res = findall(mask, info, IGNORECASE)
        return res[0] if res else ''

    def Gzh_t1_t2_creation() -> tuple[float, float, float]:
        """На основе информации из бланка возвращает три параметра

        Returns:
            tuple[float, float, float]: Gzh, t1, t2
        """
        Gzh_t1_t2 = findall(r'.*?[GV]ж=\s*?(\d+[.,]?\d*?)\D.*?tжн\*?=\s*?(-?\d+[.,]?\d*?)\D.*?tжк\*?=\s*?(-?\d+[.,]?\d*?)\D.*', info.replace(',', '.').replace('\n', ';'), IGNORECASE)
        if Gzh_t1_t2:
            Gzh, t1, t2 = map(float, Gzh_t1_t2[0])
        else:
            Gzh, t1, t2 = (
                float(''.join(takewhile(lambda x: x.isdigit(), parapa.split('=')[1]))) 
                for parapa in 
                (hehehe for hehehe in info.split('; ') for j in ('Gж=', 'tжн*=', 'tжк*=') if j in hehehe))
        return Gzh, t1, t2

    if debug_mode_on:
        pdmo.debug_mode_tumbler()

    two_streams = '/' in main_blank.main_information['Поток']
    pdmo(two_streams, main_blank.main_information['Поток'], main_blank.main_information['Название'])

    RESULT_STRING_TEMPLATE = {
        'Обозначение системы' : main_blank.main_information['Название'],
        'Кол. систем' : '',
        'Наименование обслуживаемого помещения (технологического оборудования)' : '',
        'Тип (наименование)' : main_blank.main_information['Типоразмер'] if main_blank.IS_VEROSA or main_blank.IS_KHOLOD else '',

        #     'Вентилятор',
        'Вентилятор_Исполнение по взрывозащите' : '',
        'Вентилятор_L, м3/ч' : '',
        'Вентилятор_P, Па' : '',
        'Вентилятор_n, мин-1' : '',
        'Вентилятор_Электродвигатель_Тип (наименование)' : '',
        'Вентилятор_Электродвигатель_N, кВт' : '',
        'Вентилятор_Электродвигатель_n, мин-1' : '',

        #     'Воздухонагреватель',
        'Воздухонагреватель_Тип (наименование)' : '',
        'Воздухонагреватель_Кол.' : '',
        'Воздухонагреватель_Т-ра нагрева, C_от' : '',
        'Воздухонагреватель_Т-ра нагрева, C_до' : '',
        'Воздухонагреватель_Расход теплоты, Вт' : '',
        'Воздухонагреватель_ΔP, Па_по воздуху' : '',
        'Воздухонагреватель_ΔP, Па_по воде' : '',

        #     'Рекуператор',
        'Рекуператор_Тип (наименование)' : '',
        'Рекуператор_Кол.' : '',
        'Рекуператор_Расход воздуха, м3/ч_греющий' : '',
        'Рекуператор_Расход воздуха, м3/ч_нагреваемый' : '',
        'Рекуператор_Т-ра нагрева, C_от' : '',
        'Рекуператор_Т-ра нагрева, C_до' : '',
        'Рекуператор_Расход теплоты, Вт' : '',
        'Рекуператор_n, %' : '',
        'Рекуператор_ΔP, Па_греющий' : '',
        'Рекуператор_ΔP, Па_нагреваемый' : '',

        #     'Фильтр',
        'Фильтр_Тип (наименование)' : '',
        'Фильтр_Кол.' : '',
        'Фильтр_ΔP (чистого), Па' : '',

        #     'Воздухоохладитель',
        'Воздухоохладитель_Тип (наименование)' : '',
        'Воздухоохладитель_Кол.' : '',
        'Воздухоохладитель_Т-ра охлаждения, C_от' : '',
        'Воздухоохладитель_Т-ра охлаждения, C_до' : '',
        'Воздухоохладитель_Расход холода, Вт' : '',
        'Воздухоохладитель_ΔP, Па' : '',

        #     'Насос',
        'Насос_Тип' : '',
        'Насос_G, м3/ч' : '',
        'Насос_P, МПа' : '',
        'Насос_Электродвигатель_Тип' : '',
        'Насос_Электродвигатель_N, кВт' : '',
        'Насос_Электродвигатель_n, мин-1' : '',

        'Примечание' : main_blank.main_information['Бланк-заказ'],
    }

    # result_string_2 = {key : value for key, value in result_string_1.items()}
    result_string_2 = deepcopy(RESULT_STRING_TEMPLATE)
    mult_1000 = lambda x:str(int(float(x) * 1000))
    venti_count, heata_count, cola_count = 0, 0, 0

    to_append = []

    if main_blank.IS_OTHERS:
        return to_append

    elif main_blank.IS_INDUST or main_blank.IS_INTEPU:
        return to_append

    elif main_blank.IS_OBPROM:
        pdmo(main_blank.ALL_MAIN_INFO)
        header, info = tuple(main_blank.ALL_MAIN_INFO.items())[0]
        result_temp = deepcopy(RESULT_STRING_TEMPLATE)
        
        result_temp['Тип (наименование)'] = header.split('. ')[1]
        result_temp['Вентилятор_L, м3/ч'] = findall_short(r'Q\*?=(.+?)м3\/ч')
        result_temp['Вентилятор_n, мин-1'] = findall_short(r'nдв=(.+?)[ом][би][\/н][м\-][и1]н?')
        result_temp['Вентилятор_Электродвигатель_Тип (наименование)'] = findall_short(r'назв: (.+?);')
        result_temp['Вентилятор_Электродвигатель_n, мин-1'] = findall_short(r'nрк=(.+?)[ом][би][\/н][м\-][и1]н?')
        result_temp['Вентилятор_Электродвигатель_N, кВт'] = findall_short(r'Ny=(.+?)кВт')
        all_dp = {key : float(value) for key, value in findall(r'dp(.+?)=(.+?)Па', tuple(main_blank.ALL_MAIN_INFO.items())[0][1], IGNORECASE)}                  
        result_temp['Вентилятор_P, Па'] = sum(all_dp.values()) if 'сум' not in all_dp.keys() else all_dp['сум']

        to_append.append(result_temp)

    elif main_blank.IS_KHOLOD:
        # испаритель - охладитель, конденсатор - нагреватель, компрессор - вентилятор

        info = '; '.join(main_blank.ALL_MAIN_INFO['Технические характеристики оборудования'])
        kW = mult_1000(findall_short(r'Холодопроизводительность \/ Cooling capacity; кВт \/ kW; (.+?);').replace(',', '.'))
        pdmo(info)
        if 'Испаритель / Evaporator' in info:
            result_temp = deepcopy(RESULT_STRING_TEMPLATE)
            result_temp['Воздухоохладитель_Расход холода, Вт'] = kW
            result_temp['Воздухоохладитель_Кол.'] = findall_short(r'Испаритель.+?Количество испарителей \/ Number of evaporators; шт. \/ pcs; (.+?);')
            result_temp['Воздухоохладитель_Тип (наименование)'] = findall_short(r'Испаритель.+?Тип испарителя \/ Evaporator type; (.+?) \/')
            result_temp['Воздухоохладитель_Т-ра охлаждения, C_от'] = findall_short(r'Испаритель.+?Температура теплоносителя в испарителе вход \/ Fluid temperature IN; °С; (.+?);')
            result_temp['Воздухоохладитель_Т-ра охлаждения, C_до'] = findall_short(r'Испаритель.+?Температура теплоносителя в испарителе выход \/ Fluid temperature OUT; °С; (.+?);')
            result_temp['Воздухоохладитель_ΔP, Па'] = mult_1000(findall_short(r'Испаритель.+?Гидравлическое сопротивление \/ Pressure drop; кПа \/ kPa; (.+?);').replace(',', '.'))
            
            to_append.append(result_temp)
            pass

        if 'Конденсатор / Condenser' in info:
            result_temp = deepcopy(RESULT_STRING_TEMPLATE)
            result_temp['Воздухонагреватель_Расход теплоты, Вт'] = kW
            result_temp['Воздухонагреватель_Кол.'] = findall_short(r'Конденсатор.+?Количество конденсаторов \/ Number of condensers; шт. \/ pcs; (.+?);')
            result_temp['Воздухонагреватель_Тип (наименование)'] = findall_short(r'Конденсатор.+?Тип конденсатора \/ Condenser type; (.+?) \/')
            result_temp['Воздухонагреватель_Т-ра нагрева, C_от'] = findall_short(r'Конденсатор.+?Температура теплоносителя в конденсаторе вход \/ Fluid temperature IN; °С; (.+?);')
            result_temp['Воздухонагреватель_Т-ра нагрева, C_до'] = findall_short(r'Конденсатор.+?Температура теплоносителя в конденсаторе выход \/ Fluid temperature OUT; °С; (.+?);')
            result_temp['Воздухонагреватель_ΔP, Па_по воде'] = mult_1000(findall_short(r'Конденсатор.+?Гидравлическое сопротивление \/ Pressure drop; кПа \/ kPa; (.+?);').replace(',', '.'))

            to_append.append(result_temp)
            pass

        if 'Электропитание / Power supply' in info:
            result_temp = deepcopy(RESULT_STRING_TEMPLATE)
            if 'Компрессоры / Compressors' in info:
                result_temp['Вентилятор_Исполнение по взрывозащите'] = findall_short(r'Компрессоры.+?Тип компрессора \/ Compressor type; (.+?) \/')  + ' (' + findall_short(r'Количество компрессоров \/ Number of compressors; шт\. \/ pcs; (.+?);') + ' шт)'
            result_temp['Вентилятор_Электродвигатель_Тип (наименование)'] = findall_short(r'Параметры электропитания \/ Power supply; \/Гц\/В; phi\/Hz\/V; (.+?) \/') + ' ' + findall_short(r'Общий рабочий ток \/ Total operating current; А; (.+?);').replace(',', '.') + 'А'
            result_temp['Вентилятор_Электродвигатель_N, кВт'] = findall_short(r'Общая потребляемая мощность \/ Total absorbed power; кВт \/ kW; (.+?);').replace(',', '.')

            to_append.append(result_temp)
            pass
        pass

    elif main_blank.IS_DRYDOL:
        # мощность 703,5, расход по воздуху под напряжением питания 203200 (это драйкулеры), ещё т жидкости на хводе и выходе. Драйкулеры - вентилятор и охладитель

        info = '; '.join(main_blank.ALL_MAIN_INFO['Технические характеристики оборудования'])
        pdmo(info)
        pass

    elif main_blank.IS_VEROSA:
        for header, info in main_blank.ALL_MAIN_INFO.items():
            result_temp = deepcopy(RESULT_STRING_TEMPLATE)

            if 'Дополнительное оборудование' in header:
                continue
            if main_blank.IS_CHANAL:
                info = info.replace(',', '.')

            if 'вентилятор' in header.lower() or 'вентилятор' in info.lower():
                pdmo(header, info)
                if 'блок перехода на' in header.lower():
                    continue

                if 'камера' in header.lower():
                    continue

                if main_blank.IS_VEROSA:
                    res_ven_Q = findall(r'Q\*?=(.+?)м3\/ч', info, IGNORECASE)
                    if res_ven_Q:
                        res_ven_Q = res_ven_Q[0]
                    else:
                        for rere in main_blank.docx_text:
                            if 'Lв=' in rere:
                                res_ven_Q = findall(r'Lв=(.+?)м', rere, IGNORECASE)[0]
                                pdmo(res_ven_Q)
                                break
                            if res_ven_Q:
                                break
                    
                    result_temp['Вентилятор_L, м3/ч'] = res_ven_Q

                    all_dp = {key : float(value) for key, value in findall(r'dp(.+?)=(.+?)Па', info, IGNORECASE)}

                    # result_string_1['Вентилятор_L, м3/ч'] = findall_short(r'Q\*?=(.+?)м3\/ч')
                    result_temp['Вентилятор_n, мин-1'] = findall_short(r'nдв=(.+?)[ом][би][\/н][м\-][и1]н?')
                    result_temp['Вентилятор_Электродвигатель_Тип (наименование)'] = findall_short(r'назв: (.+?);')
                    result_temp['Вентилятор_Электродвигатель_n, мин-1'] = findall_short(r'nрк=(.+?)[ом][би][\/н][м\-][и1]н?')
                    result_temp['Вентилятор_Электродвигатель_N, кВт'] = findall_short(r'Ny=(.+?)кВт')
                    result_temp['Вентилятор_P, Па'] = sum(all_dp.values()) if 'сум' not in all_dp.keys() else all_dp['сум']

                    venti_count += 1

                    to_append.append(result_temp)

                elif main_blank.IS_CHANAL:
                    pdmo(header, info)

                    result_temp['Тип (наименование)'] = findall_short(r'Индекс: (.+?);')

                    result_temp['Вентилятор_L, м3/ч'] = findall_short(r'Lв=(.+?) куб')
                    # result_temp['Вентилятор_n, мин-1'] = findall_short(r'')
                    result_temp['Вентилятор_P, Па'] = findall_short(r'Pполн=(.+?) Па')
                    # result_temp['Вентилятор_Электродвигатель_Тип (наименование)'] = findall_short(r'')
                    # result_temp['Вентилятор_Электродвигатель_n, мин-1'] = findall_short(r'')
                    result_temp['Вентилятор_Электродвигатель_N, кВт'] = findall_short(r'Ny=(.+?) кВт;')

                    to_append.append(result_temp)

                    pass
                pass

            elif 'нагрев' in header.lower() or 'нагрев' in info.lower():
                pdmo(header, info)
                result_temp['Воздухонагреватель_Кол.'] = '1'

                if 'жидкостный' in header.lower() or 'подготовки воздуха,нагрев(вода)' in header.lower() or 'водяной' in header.lower():
                    try:
                        heater_info = findall(r'Qт=(.+?)кВт.+?tвн\*?=(.+?)°C; tвк\*?=(.+?)°C.+?dpво\*?=(.+?)Па.+?dpж\*?=(.+?)кПа' if main_blank.IS_VEROSA else r'Qt=(.+?)кВт.+?tвн\*?=(.+?)°C; tвк\*?=(.+?)°C.+?dPж=(.+?)кПа.+?dPв=(.+?)Па', info, IGNORECASE)[0]
                    except IndexError:
                        pdmo(header, info)
                    else:
                        Gzh, t1, t2 = Gzh_t1_t2_creation()
                        # qweqweqwe = vector_creation(Gzh, 'С' if t1 >= 100 else ('Ш' if Gzh <= 14_000 else 'С'), 2)
                        # pdmo(qweqweqwe)

                        result_temp['Воздухонагреватель_Тип (наименование)'] = 'жидкостный'
                        if main_blank.IS_VEROSA:
                            result_temp['Воздухонагреватель_Расход теплоты, Вт'], result_temp['Воздухонагреватель_Т-ра нагрева, C_от'], result_temp['Воздухонагреватель_Т-ра нагрева, C_до'], result_temp['Воздухонагреватель_ΔP, Па_по воздуху'], result_temp['Воздухонагреватель_ΔP, Па_по воде'] = heater_info
                        else:
                            result_temp['Воздухонагреватель_Расход теплоты, Вт'], result_temp['Воздухонагреватель_Т-ра нагрева, C_от'], result_temp['Воздухонагреватель_Т-ра нагрева, C_до'], result_temp['Воздухонагреватель_ΔP, Па_по воде'], result_temp['Воздухонагреватель_ΔP, Па_по воздуху'] = heater_info
                        result_temp['Воздухонагреватель_Расход теплоты, Вт'] = mult_1000(result_temp['Воздухонагреватель_Расход теплоты, Вт'])
                        result_temp['Воздухонагреватель_ΔP, Па_по воде'] = mult_1000(result_temp['Воздухонагреватель_ΔP, Па_по воде'])

                        result_temp['Насос_G, м3/ч'] = str(int(Gzh))
                        # result_temp['Насос_Тип'] = qweqweqwe[4]
                        result_temp['Насос_P, МПа'] = str(round(float(result_temp['Воздухонагреватель_ΔP, Па_по воде']) / 1_000_000, 3))

                        to_append.append(result_temp)
                        pass
                elif 'электрический' in header.lower() or 'подготовки воздуха,нагрев(эл)' in header.lower():
                    try:
                        heater_info = findall(r'Q[тt]=(.+?)кВт.+?tвн\*?=(.+?)°C; tвк\*?=(.+?)°C.+?d[Pp]во?\*?=(.+?)Па', info, IGNORECASE)[0]
                    except IndexError:
                        pdmo(header, info)
                        pass
                    else:
                        result_temp['Воздухонагреватель_Тип (наименование)'] = 'электрический'

                        result_temp['Воздухонагреватель_Расход теплоты, Вт'], result_temp['Воздухонагреватель_Т-ра нагрева, C_от'], result_temp['Воздухонагреватель_Т-ра нагрева, C_до'], result_temp['Воздухонагреватель_ΔP, Па_по воздуху'] = heater_info
                        result_temp['Воздухонагреватель_Расход теплоты, Вт'] = mult_1000(result_temp['Воздухонагреватель_Расход теплоты, Вт'])

                        to_append.append(result_temp)
                        pass
                elif 'паровой' in header.lower():
                    try:
                        heater_info = findall(r'Qт=(.+?)кВт.+?tвн\*?=(.+?)°C; tвк\*?=(.+?)°C.+?dpво\*?=(.+?)Па', info, IGNORECASE)[0]
                    except IndexError:
                        pdmo(header, info)
                        pass
                    else:
                        result_temp['Воздухонагреватель_Тип (наименование)'] = 'паровой'

                        result_temp['Воздухонагреватель_Расход теплоты, Вт'], result_temp['Воздухонагреватель_Т-ра нагрева, C_от'], result_temp['Воздухонагреватель_Т-ра нагрева, C_до'], result_temp['Воздухонагреватель_ΔP, Па_по воздуху'] = heater_info
                        result_temp['Воздухонагреватель_Расход теплоты, Вт'] = mult_1000(result_temp['Воздухонагреватель_Расход теплоты, Вт'])

                        to_append.append(result_temp)
                        
                    pass
                elif 'газовый' in header.lower():
                    result_temp['Воздухонагреватель_Тип (наименование)'] = 'газовый'
                    to_append.append(result_temp)
                    pass
                elif 'дизельный' in header.lower():
                    result_temp['Воздухонагреватель_Тип (наименование)'] = 'дизельный'
                    to_append.append(result_temp)
                    pass

                # result_string_2 = {key : value if 'Воздухонагреватель' in key else result_string_2[key] for key, value in result_string_1.items()}
                # result_string_2 = {key : value for key, value in result_string_1.items()}
                pass

            elif 'теплоутилизатор' in header.lower() or 'теплоутилизатор' in info.lower():
                pdmo(header, info)
                if 'охладитель с промеж.теплоносителем' in header.lower():
                    continue
                
                result_temp_2 = deepcopy(RESULT_STRING_TEMPLATE)

                if main_blank.IS_VEROSA:
                    result_temp['Рекуператор_Кол.'], result_temp_2['Рекуператор_Кол.'] = 1, 1
                    result_temp['Рекуператор_Тип (наименование)'], result_temp_2['Рекуператор_Тип (наименование)'] = ('роторный' if 'ротор' in header.lower() else 'пластинчатый' if 'пластинчатый' in header.lower() else 'жидкостный' if 'нагреватель' in header.lower() else '', ) * 2
                    result_temp['Рекуператор_Расход теплоты, Вт'], result_temp_2['Рекуператор_Расход теплоты, Вт'] = (mult_1000(findall_short(r'Qп=(.+?)кВт')), ) * 2
                    result_temp['Рекуператор_Расход воздуха, м3/ч_греющий'] = findall_short(r'Lв0п=(.+?)м3/ч')
                    result_temp['Рекуператор_Расход воздуха, м3/ч_нагреваемый'] = findall_short(r'Lвкп=(.+?)м3/ч')
                    result_temp_2['Рекуператор_Расход воздуха, м3/ч_греющий'] = findall_short(r'Lв0в=(.+?)м3/ч')
                    result_temp_2['Рекуператор_Расход воздуха, м3/ч_нагреваемый'] = findall_short(r'Lвкв=(.+?)м3/ч')
                    result_temp['Рекуператор_ΔP, Па_греющий'], result_temp_2['Рекуператор_ΔP, Па_греющий'] = (findall_short(r'dpв0?п=(.+?)Па'), ) * 2
                    result_temp['Рекуператор_ΔP, Па_нагреваемый'], result_temp_2['Рекуператор_ΔP, Па_нагреваемый'] = (findall_short(r'dpв0?в=(.+?)Па'), ) * 2
                    result_temp['Рекуператор_n, %'] = findall_short(r'КПДсп=(.+?)%' if 'КПДсп' in info else r'Ktп=(.+?)%')
                    result_temp_2['Рекуператор_n, %'] = findall_short(r'КПДсв=(.+?)%' if 'КПДсв' in info else r'Ktв=(.+?)%')
                    result_temp['Рекуператор_Т-ра нагрева, C_от'] = findall_short(r'tвнп=(.+?)°C')
                    result_temp['Рекуператор_Т-ра нагрева, C_до'] = findall_short(r'tвкп=(.+?)°C')
                    result_temp_2['Рекуператор_Т-ра нагрева, C_от'] = findall_short(r'tвнв=(.+?)°C')
                    result_temp_2['Рекуператор_Т-ра нагрева, C_до'] = findall_short(r'tвкв=(.+?)°C')

                    pdmo(result_temp, result_temp_2)
                    pass

                elif main_blank.IS_CHANAL:
                    recup_info = list(findall(r'КПД=(.+?) %; tвн_приток=(.+?) °C; tвк_приток=(.+?) °C; tвн_вытяжка=(.+?) °C; tвк_вытяжка=(.+?) °C; dPв_приток=(.+?) Па; dPв_вытяжка=(.+?) Па', info, IGNORECASE)[0])
                    result_temp['Рекуператор_Тип (наименование)'], result_temp_2['Рекуператор_Тип (наименование)'] = ('роторный' if 'ротор' in header.lower() else 'пластинчатый' if 'пластинчатый' in header.lower() else 'жидкостный' if 'нагреватель' in header.lower() else '', ) * 2
                    result_temp['Рекуператор_n, %'], result_temp_2['Рекуператор_n, %'] = (recup_info.pop(0), ) * 2
                    result_temp['Рекуператор_ΔP, Па_греющий'], result_temp_2['Рекуператор_ΔP, Па_греющий'], result_temp['Рекуператор_ΔP, Па_нагреваемый'], result_temp_2['Рекуператор_ΔP, Па_нагреваемый'] = (recup_info.pop(-2), recup_info.pop(-1)) * 2
                    result_temp['Рекуператор_Т-ра нагрева, C_от'], result_temp['Рекуператор_Т-ра нагрева, C_до'], result_temp_2['Рекуператор_Т-ра нагрева, C_от'], result_temp_2['Рекуператор_Т-ра нагрева, C_до'] = recup_info
                    pass

                to_append += [result_temp, result_temp_2]

            elif 'фильтр' in header.lower() or 'фильтр' in info.lower():
                pdmo(info)
                result_temp['Фильтр_Кол.'] = '1'

                # filter_info = findall(r'класс: (.+?);.+?dpвр=(.+?)Па', info, IGNORECASE)[0]
                if main_blank.IS_VEROSA:
                    result_temp['Фильтр_Тип (наименование)'] = findall_short(r'класс: (.+?);')
                    result_temp['Фильтр_ΔP (чистого), Па'] = findall_short(r'dpв=(.+?)Па')
                elif main_blank.IS_CHANAL:
                    filter_info = findall(r'Класс: (.+?);.+?dPв=(.+?)\s?Па;', info, IGNORECASE)[0]
                    result_temp['Фильтр_Тип (наименование)'], result_temp['Фильтр_ΔP (чистого), Па'] = filter_info
                
                to_append.append(result_temp)

            elif 'воздухоохладитель' in header.lower() or 'воздухоохладитель' in info.lower():
                pdmo(info)
                try:
                    cooler_info = findall(r'Qх=(.+?)кВт;.+?tвн=(.+?)°C;.+?tвк=(.+?)°C;.+?dpво=(.+?)Па;', info, IGNORECASE)[0]
                except IndexError:
                    pdmo(header, info)
                    pass
                else:
                    result_temp['Воздухоохладитель_Кол.'] = '1'
                    result_temp['Воздухоохладитель_Тип (наименование)'] = 'жидкостный' if 'жидкостный' in header.lower() else 'фреоновый'
                    result_temp['Воздухоохладитель_Расход холода, Вт'], result_temp['Воздухоохладитель_Т-ра охлаждения, C_от'], result_temp['Воздухоохладитель_Т-ра охлаждения, C_до'], result_temp['Воздухоохладитель_ΔP, Па'] = cooler_info
                    result_temp['Воздухоохладитель_Расход холода, Вт'] = mult_1000(result_temp['Воздухоохладитель_Расход холода, Вт'])
                    to_append.append(result_temp)
                    pass
                pass
            pass

    else:
        pass
    pdmo(to_append)

    if main_blank.IS_CHANAL:
        name_name = tuple(el['Тип (наименование)'] for el in to_append if el['Тип (наименование)'])
        if name_name:
            pdmo(name_name)
            pass
            for i in range(len(to_append)):
                to_append[i]['Тип (наименование)'] = name_name[0]

    if debug_mode_on:
        pdmo.debug_mode_tumbler()
    return to_append

def sort_rule(the_result:list):
    """_summary_

    Args:
        the_result (list): _description_

    Returns:
        _type_: _description_
    """
    
    # И порядок должен быть:
    # 1. Системы К
    # 2. Системы П
    # 3. Системы ПВ
    # 4. Системы В
    # 5. Системы ДВ
    # 6. Системы ДП

    sys_K, sys_P, sys_PV, sys_V, sys_DV, sys_DP, sys_other = [], [], [], [], [], [], []
    # r'^П.*?\/?В'
    for el in the_result:
        name = el['Обозначение системы']
        if not ('ДП' in name.upper() or 'ПД' in name.upper()):
            if not ('ДВ' in name.upper() or 'ВД' in name.upper() or 'ДУ' in name.upper()):
                if name[0] == 'К':
                    sys_K.append(el)
                else:
                    if 'П' in name.upper():
                        if not findall(r'^П.*?\/?В', name, IGNORECASE):
                            sys_P.append(el)
                        else:
                            sys_PV.append(el)
                    else:
                        if 'В' in name.upper():
                            sys_V.append(el)
                        else:
                            sys_other.append(el)
            else:
                sys_DV.append(el)
        else:
            sys_DP.append(el)
    pdmo(sys_K, sys_P, sys_PV, sys_V, sys_DV, sys_DP, sys_other)
    return sys_K + sys_P + sys_PV + sys_V + sys_DV + sys_DP + sys_other

if __name__ == "__main__":
    # print(os.listdir(os.path.join(os.environ['WINDIR'],'fonts')))
    # print(os.listdir(r'C:\Windows\fonts'))
    debug_mode_on = True
    result = []

    test_files = tuple(Path(file) for file in (
        'e:/Python/21_AUTOMATA/_ТЗ_для программы/Веросы/Бланк1/211027853-ОПР АЭРОПРОЕКТ (Реконстр. аэропортового комплекса Чертовицкое г.Воронеж ОАСС).docx',  # 0
        'e:/Python/21_AUTOMATA/_ТЗ_для программы/Веросы/231038867-СПБ_ПВ4е.doc',  # 1
        'e:/Python/21_AUTOMATA/_ТЗ_для программы/Веросы/231036475-СПБ.doc',  # 2
        'e:/Python/21_AUTOMATA/_ТЗ_для программы/Веросы/231036476-СПБ.doc',  # 3
        'e:/Python/21_AUTOMATA/_ТЗ_для программы/Веросы/231036477-СПБ.doc',  # 4
        'e:/Python/21_AUTOMATA/_ТЗ_для программы/Веросы/231036478-СПБ.doc',  # 5
        'e:/Python/21_AUTOMATA/_ТЗ_для программы/Веросы/231036736-СПБ_AHU1 версия 2.doc',  # 6
        'ТЕСТОВЫЕ ФАЙЛЫ/ПВ СИСТЕМЫ/231040740а-ОПР ПВ14 ДИ БИ СИ (Аэропорт Хабаровск, МВЛ).doc',  # 7
        'ТЕСТОВЫЕ ФАЙЛЫ/231058201-ОПР ПВ4 ГИПРОЗДРАВ (Клиника здоровья, г.о. Истра п. Пансионат Березка, 3у 1).doc',  # 8
        'ТЕСТОВЫЕ ФАЙЛЫ/Каналка/П4,В6.docx',  # 9
        'Новая папка/234100449а-ОПР МЕСА ИНЖИНИРИНГ (Жилой м.р-н для Амурского ГПЗ в г.Свободный)_ТЕХ.doc',  # 10
        'Новая папка/МАВО.D.41.900.2х4.А.6R.pdf',  # 11
    ))
    
    # pdf_reader = pypdf.PdfReader(test_files[11])
    # for page in pdf_reader.pages:
    #     pdmo(page.extract_text())
    # pdmo(pdf_reader.pages[0].extract_text())

    pdmo.debug_mode_tumbler()
    blank = Blank(test_files[11])
    pdmo.debug_mode_tumbler()

    the_test_file = hovs(blank)
    print(the_test_file)
    pass

    for test_file in test_files:
        if not test_file.exists():
            continue
        file_result = hovs(Blank(test_file))
        pdmo(file_result)
        result += file_result

    result = sort_rule(result)
    pdmo(result)
        
    Path('hovs.xlsx').write_bytes(two_excel_calc(DataFrame(result, columns=result[0].keys())).getbuffer())
else:
    debug_mode_on = False