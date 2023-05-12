from tkinter import *
import tkinter as tk
import time
import sys
from tkinter import filedialog as fd
import openpyxl
import numpy as np
import math
from termcolor import colored


rasstoyanie_mejdu_skv =[]
ogranicheniya = []
parametri = []


def rasschet_final():
#    text_box.delete('1.0', 'end')
    text_box.insert(END, 'Входные данные:\n')
# ВЫВОД ТИПА МЕСТОРОЖДЕНИЯ В КОНСОЛЬ
    if var.get() == 0:
        FieldType = 'Газовое'
    elif var.get() == 1:
        FieldType = 'Газоконденсатное'
    else:
        FieldType = 'Нефтяное'
    text_box.insert(END, 'Тип месторождения - ' + FieldType)

    if int(intvar_cb38.get()) == 0:
        text_box.insert(END, '\nПодошвенная вода отсутствует')
    else:
        text_box.insert(END, '\nНаличие подошвенной воды')

#ВЫВОД ОГРАНИЧЕНИЙ В КОНСОЛЬ

    if intvar_cb4.get() == 0:
        pass
    else:
        ogranicheniya.append('Максимальный безводный дебит')

    if intvar_cb5.get() == 0:
        pass
    else:
        ogranicheniya.append('Минимальный дебит для выноса воды')

    if intvar_cb6.get() == 0:
        pass
    else:
        ogranicheniya.append('Минимальный дебит для выноса конденсата')

    if intvar_cb7.get() == 0:
        pass
    else:
        ogranicheniya.append('Ограничение скорости потока для уменьшение коррозии')

    if intvar_cb8.get() == 0:
        pass
    else:
        ogranicheniya.append('Гидратообразование на забое скважины')

    if intvar_cb9.get() == 0:
        pass
    else:
        ogranicheniya.append('Гидратообразование на устье скважины')

    if intvar_cb10.get() == 0:
        pass
    else:
        ogranicheniya.append('Гидратообразование в стволе скважины')

    if intvar_cb11.get() == 0:
        pass
    else:
        ogranicheniya.append('Гидратообразование в шлейфе')

    if intvar_cb12.get() == 0:
        pass
    else:
        ogranicheniya.append('Ограничение скорости потока в шлейфе')

    text_box.insert(END, '\n\nУчитываемые ограничения:\n')
    for i in range(0, len(ogranicheniya)):
        text_box.insert(END, '\n')
        text_box.insert(END, ogranicheniya[i])

#ВЫВОД ПАРАМЕТРОВ В КОНСОЛЬ
    active_parameters_letters =['C']
    if intvar_cb14.get() == 0:
        pass
    else:
        parametri.append('Статическое давление скважины')
        active_parameters_letters.append('D')

    if intvar_cb15.get() == 0:
        pass
    else:
        parametri.append('Пластовое давление')
        active_parameters_letters.append('E')

    if intvar_cb16.get() == 0:
        pass
    else:
        parametri.append('Рабочий дебит скважины')
        active_parameters_letters.append('F')

    if intvar_cb17.get() == 0:
        pass
    else:
        parametri.append('Депрессия скважины')
        active_parameters_letters.append('G')

    if intvar_cb18.get() == 0:
        pass
    else:
        parametri.append('Устьевое давление скважины')
        active_parameters_letters.append('H')

    if intvar_cb19.get() == 0:
        pass
    else:
        parametri.append('Затрубное давление скважины')
        active_parameters_letters.append('I')

    if intvar_cb20.get() == 0:
        pass
    else:
        parametri.append('Давление в шлейфе')
        active_parameters_letters.append('J')

    if intvar_cb21.get() == 0:
        pass
    else:
        parametri.append('Устьевая температура скважины')
        active_parameters_letters.append('K')

    if intvar_cb22.get() == 0:
        pass
    else:
        parametri.append('Межколонное давление')
        active_parameters_letters.append('L')

    if intvar_cb23.get() == 0:
        pass
    else:
        parametri.append('Давление на входе в УКПГ')
        active_parameters_letters.append('M')

    if intvar_cb24.get() == 0:
        pass
    else:
        parametri.append('Температура на входе в УКПГ')
        active_parameters_letters.append('N')

    if intvar_cb25.get() == 0:
        pass
    else:
        parametri.append('Дебит скважины по воде')
        active_parameters_letters.append('O')

    text_box.insert(END, '\n\nУчитываемые параметры:\n')
    for i in range(0, len(parametri)):
        text_box.insert(END, '\n')
        text_box.insert(END, parametri[i])

#ВЫВОД СКВАЖИН В КОНСОЛЬ
    skv = int(Entry1.get())
    text_box.insert(END, '\n\nКоличество скважин - ')
    if skv > 16:
        skv = 16
    else:
        pass
    text_box.insert(END, str(skv) + '\n')
    for j in range(0, skv-1):
        text_box.insert(END, '\nРасстояние между ' + str(j+1) + '-й и ' + str(j+2) + '-й скважинами '+ Entries[j].get() + ' метров')
        rasstoyanie_mejdu_skv.append(int(Entries[j].get()))

    print(rasstoyanie_mejdu_skv)

    text_box.insert(END, '\n\nКоличество итераций - ' + str(Entry17.get() + ' итераций'))
# ВЫВОД ПРОГНОЗОВ В КОНСОЛЬ
    prognozi = []
    text_box.insert(END, '\n\nПрогнозируемые параметры:\n')
    if intvar_cb27.get() == 0:
        pass
    else:
        prognozi.append('Дебит накопленный по максимально допустимому режиму на ' + Entry18.get() + ' месяцев')
    if intvar_cb28.get() == 0:
        pass
    else:
        prognozi.append('Дебит накопленный по минимально допустимому режиму на '  + Entry18.get() + ' месяцев')
    if intvar_cb29.get() == 0:
        pass
    else:
        prognozi.append('Дебит накопленный по оптимальному режиму на ' + Entry18.get() + ' месяцев')
    if intvar_cb30.get() == 0:
        pass
    else:
        prognozi.append('Изменение пластового давления Pпл при максимально допустимом режиме на ' + Entry18.get() + ' месяцев')
    if intvar_cb31.get() == 0:
        pass
    else:
        prognozi.append('Изменение пластового давления Pпл при минимально допустимом режиме на ' + Entry18.get() + ' месяцев')
    if intvar_cb32.get() == 0:
        pass
    else:
        prognozi.append('Изменение пластового давления Pпл при оптимальном режиме на ' + Entry18.get() + ' месяцев')
    if intvar_cb33.get() == 0:
        pass
    else:
        prognozi.append('Изменение устьевого давления Pу при максимально допустимом режиме на ' + Entry18.get() + ' месяцев')
    if intvar_cb34.get() == 0:
        pass
    else:
        prognozi.append('Изменение устьевого давления Pу при минимально допустимом режиме на ' + Entry18.get() + ' месяцев')
    if intvar_cb35.get() == 0:
        pass
    else:
        prognozi.append('Изменение устьевого давления Pу при оптимальном режиме на ' + Entry18.get() + ' месяцев')
    for i in range(0, len(prognozi)):
        text_box.insert(END, '\n')
        text_box.insert(END, prognozi[i])
    file_name = fd.askopenfilename()
    wb = openpyxl.reader.excel.load_workbook(filename=file_name)
    wb.active = 0
    wb.active2 = 1
    sheet = wb.active
    sheet2 = wb.active2
    training_inputs2 = [[sheet['C4'].value,sheet['D4'].value, sheet['E4'].value, sheet['F4'].value, sheet['G4'].value,
                         sheet['H4'].value, sheet['I4'].value, sheet['J4'].value, sheet['K4'].value,
                         sheet['L4'].value, sheet['M4'].value, sheet['N4'].value, sheet['O4'].value]]
    for i in range(5, 50):
        training_inputs2.append([sheet['C' + str(i)].value, sheet['D' + str(i)].value, sheet['E' + str(i)].value, sheet['F' + str(i)].value,
                                 sheet['G' + str(i)].value, sheet['H' + str(i)].value, sheet['I' + str(i)].value,
                                 sheet['J' + str(i)].value, sheet['K' + str(i)].value, sheet['L' + str(i)].value,
                                 sheet['M' + str(i)].value, sheet['N' + str(i)].value, sheet['O' + str(i)].value])

    letters = ['C','D','E','F','G','H','I','J','K','L','M','N','O']
    input_names_colored = ['Дата', 'Pст', 'Рпл', 'Qж', 'Депр','Рус', 'Рзт', 'Ршл',
                  'Тус', 'Рмк','Рвх', 'Твх', 'Qв']
    reds = []
    whites =[]
    text_box.insert(END, '\n')
    text_box.tag_config('tag_red_text', foreground='red')
    text_box.insert(END, '\nКрасный - параметр не используется при расчетах весов, только при необходимости для '
                         'вычисления минимальных/макисимальных    границ дебита/давления/скорости ограничений\n', 'tag_red_text')
    text_box.insert(END, '\n')
    for i in range(0, len(letters)):
        if letters[i] in active_parameters_letters:
            whites.append(i)
            if i == 0:
                text_box.insert(END, str(input_names_colored[i]))
                text_box.insert(END, '                ')
            else:
                text_box.insert(END, str(input_names_colored[i]))
                text_box.insert(END, '       ')
        else:
            reds.append(i)
            text = str(input_names_colored[i])
            text_box.insert(END, text, 'tag_red_text')
            text_box.insert(END,'       ')
    for j in range(4, 50):
        text_box.insert(END, '\n')
        for i in range(0, len(letters)):
            if i in reds:
                text = str(training_inputs2[j-4][i])
                text_box.insert(END, text, 'tag_red_text')
                text_box.insert(END, '      ')
            else:
                text = str(training_inputs2[j-4][i])
                text_box.insert(END, text)
                text_box.insert(END, '      ')


    Q_bezvodn_minimals = []
    def obvodnenie(i):
        h_otn = sheet['AD' + str(i)].value / sheet['AE' + str(i)].value
        R_otn = sheet['AI' + str(i)].value / sheet['AH' + str(i)].value
        az = sheet['AB' + str(i)].value * sheet['AD' + str(i)].value / math.log(R_otn)
        bz = sheet['AC' + str(i)].value * sheet['AD' + str(i)].value ** 2 / (
                    1 / sheet['AH' + str(i)].value - 1 / sheet['AI' + str(i)].value)
        deltaP_sqr = (0.1 * (sheet['AD' + str(i)].value - sheet['AE' + str(i)].value) * (
                    sheet['AF' + str(i)].value - sheet['AG' + str(i)].value)) ** 2
        K0 = 4 * bz * deltaP_sqr / (az ** 2 * sheet['AH' + str(i)].value)
        Qz = h_otn * math.log(R_otn / h_otn) * (
                    -1 + math.sqrt(1 + K0 / (h_otn * math.log(R_otn / h_otn) * math.log(R_otn / h_otn))))
        Q_bezvodn = az * Qz * sheet['AH' + str(i)].value / (2 * bz)
        Q_bezvodn_minimals.append(Q_bezvodn)
        if 'Максимальный безводный дебит' in ogranicheniya:
            text_box.insert(END, '\nМаксимальный безводный дебит = ' + str("%.3f" % Q_bezvodn) + ' тыс.м3/сут\n')

        # функция ограничения по водяной пробке

    q_water_min_minimals =[]
    def water_probka(i):
        T_critical = 125 * (1 + sheet['W' + str(i)].value)
        P_critical = 0.4903 * (10 - sheet['W' + str(i)].value)
        T_priv = (sheet['K' + str(i)].value + 273.15) / T_critical
        P_priv = sheet['H' + str(i)].value / P_critical
        Z = (0.4 * math.log10(T_priv) + 0.73) ** P_priv + 0.1 * P_priv
        v_water_min = 1.23 * (
                    45 - 0.45 * (sheet['E' + str(i)].value - sheet['G' + str(i)].value)) ** 0.25 / math.sqrt(
            0.45 * (sheet['E' + str(i)].value
                    - sheet['G' + str(i)].value))

        q_water_min = 60 * 60 * 24 / 1000 * (
                v_water_min * 273.15 * (sheet['E' + str(i)].value - sheet['G' + str(i)].value) * 3.14 *
                ((sheet['Y' + str(i)].value) / 100) ** 2) / (4 * 0.1 * Z * (sheet['X' + str(i)].value + 273.15))
        q_water_min_minimals.append(q_water_min)
        if 'Минимальный дебит для выноса воды' in ogranicheniya:
            text_box.insert(END, '\nМинимальный дебит для выноса воды = ' + str(round(q_water_min,2)) + ' тыс.м3/сут\n')
        else:
            pass


    # функция ограничения по конденсатной пробке
    q_condensat_min_minimals = []
    def condensat_probka(i):
        T_critical = 125 * (1 + sheet['W' + str(i)].value)
        P_critical = 0.4903 * (10 - sheet['W' + str(i)].value)
        T_priv = (sheet['K' + str(i)].value + 273.15) / T_critical
        P_priv = sheet['H' + str(i)].value / P_critical
        Z = (0.4 * math.log10(T_priv) + 0.73) ** P_priv + 0.1 * P_priv
        v_condensat_min = 1.71 * (67 - 0.45 * (sheet['E' + str(i)].value - sheet['G' + str(i)].value)) ** 0.25 / (
                0.45 * (sheet['E' + str(i)].value - sheet['G' + str(i)].value)) ** 0.5
        q_condensat_min = 60 * 60 * 24 / 1000 * (
                v_condensat_min * 273.15 * (sheet['E' + str(i)].value - sheet['G' + str(i)].value)
                * 3.14 * ((sheet['Y' + str(i)].value) / 100) ** 2) / (
                                    4 * 0.1 * Z * (sheet['X' + str(i)].value + 273.15))
        q_condensat_min_minimals.append(q_condensat_min)
        if 'Минимальный дебит для выноса конденсата' in ogranicheniya:
            text_box.insert(END, 'Минимальный дебит для выноса конденсата = ' + str(round(q_condensat_min,2)) + ' тыс.м3/сут')
        else:
            pass

        # функция ограничения по скорости в скважине 5-11 м/с
    vehicle = []
    q_critical_5_minimals = []
    def vehicle_well(i):
        T_critical = 125 * (1 + sheet['W' + str(i)].value)
        P_critical = 0.4903 * (10 - sheet['W' + str(i)].value)
        T_priv = (sheet['K' + str(i)].value + 273.15) / T_critical
        P_priv = sheet['H' + str(i)].value / P_critical
        Z = (0.4 * math.log10(T_priv) + 0.73) ** P_priv + 0.1 * P_priv
        q_critical_vehicle_11 = (11 * sheet['Y' + str(i)].value ** 2 * sheet['H' + str(i)].value) / (
                0.052 * (sheet['K' + str(i)].value + 273.15) * Z)
        q_critical_vehicle_5 = (5 * sheet['Y' + str(i)].value ** 2 * sheet['H' + str(i)].value) / (
                0.052 * (sheet['K' + str(i)].value + 273.15) * Z)
        q_critical_5_minimals.append(q_critical_vehicle_5)
        if sheet['F' + str(i)].value < q_critical_vehicle_11 and sheet['F' + str(i)].value > q_critical_vehicle_5:
            vehicle.append(sheet['F' + str(i)].value)
        elif sheet['F' + str(i)].value > q_critical_vehicle_11:
            if 'Ограничение скорости потока для уменьшение коррозии' in ogranicheniya:
                text_box.insert(END, '\nДебит для защиты от коррозионных элементов и выноса примесей от ' + str(round(q_critical_vehicle_5,2)) +
                    ' до ' + str(round(q_critical_vehicle_11,2)) + ' тыс. м3/сут \n')
                text_box.insert(END, 'Рекомендуется уменьшить дебит до ' + str(round(q_critical_vehicle_5,2)) + '-' + str(
                    round(q_critical_vehicle_11,2)) + ' тыс. м3/сут')
                vehicle.append(q_critical_vehicle_11)
        else:
            if 'Ограничение скорости потока для уменьшение коррозии' in ogranicheniya:
                text_box.insert(END, 'Рекомендуется увеличить дебит до ' + str(round(q_critical_vehicle_5,2)) + '-' + str(
                    round(q_critical_vehicle_11, 2)) + ' тыс. м3/сут')
                vehicle.append(q_critical_vehicle_5)

    # функция ограничения по давлению в шлейфе
    def pressure_shleif_min(x, y):
        global pipe_pressure
        pipe_pressure = sheet['J' + str(x)].value
        for i in range(x, y):
            if pipe_pressure > sheet['J' + str(i)].value:
                pipe_pressure = sheet['J' + str(i)].value
            else:
                pass
        print('МИНИМАЛЬНОЕ ДАВЛЕНИЕ В ШЛЕЙФЕ = ' + str(round(pipe_pressure,2)) + ' МПа')

    # ограничение по Pвх
    def pressure_vhod(x, y):
        for i in range(x, y):
            if sheet['M' + str(i)].value < pipe_pressure:
                pass
            else:
                print('Недостаточное давление во входе в УКПГ в скважине ' + str(i))


    #ограничение по гидратам на устье
    T_ust = []
    def hydrates_wellhead(i):
        density = sheet['W' + str(i)].value
        e = density

        if sheet['K' + str(i)].value > 0:
            b = -12420 * e ** 5 + 50909 * e ** 4 - 83004 * e ** 3 + 67296 * e ** 2 - 27145 * e + 4375.1
            T = 18.47 * math.log10(sheet['H' + str(i)].value) - b + 18.47
            T_ust.append(T)
        else:
            b1 = -39701 * e ** 5 + 162993 * e ** 4 - 266162 * e ** 3 + 216107 * e ** 2 - 87295 * e + 14090
            T = -58.51 * math.log10(sheet['H' + str(i)].value) - b1 - 58.51
            T_ust.append(T)
        if T < sheet['K' + str(i)].value:
            if 'Гидратообразование на устье скважины' in ogranicheniya:
                text_box.insert(END, '\nТекущая температура на устье скважины = ' + str(round(sheet['K' + str(i)].value,2)) + ' C\n')
                text_box.insert(END, 'Гидраты на устье образуются при температуре ниже ' + str(round(T,2)) + ' C')
        else:
            if 'Гидратообразование на устье скважины' in ogranicheniya:
                text_box.insert(END, "В скважине " + str(i) + " возможно образование гидратов\n")
                text_box.insert(END, 'Гидраты на устье образуются при температуре ниже ' + str(round(T,2)) + ' C')

    # ограничение по гидратам на забое
    T_bot = []
    def hydrates_bottom(i):
        density = sheet['W' + str(i)].value
        e = density
        if sheet['X' + str(i)].value > 0:
            b = -12420 * e ** 5 + 50909 * e ** 4 - 83004 * e ** 3 + 67296 * e ** 2 - 27145 * e + 4375.1
            T = 18.47 * math.log10(sheet['E' + str(i)].value - sheet['G' + str(i)].value) - b + 18.47
            T_bot.append(T)
        else:
            b1 = -39701 * e ** 5 + 162993 * e ** 4 - 266162 * e ** 3 + 216107 * e ** 2 - 87295 * e + 14090
            T = -58.51 * math.log10(sheet['E' + str(i)].value - sheet['G' + str(i)].value) - b1 - 58.51
            T_bot.append(T)
        if T < sheet['X' + str(i)].value:
            if 'Гидратообразование на забое скважины' in ogranicheniya:
                text_box.insert(END, '\nТекущая температура на забое скважины = ' + str(round(sheet['X' + str(i)].value,2)) + ' C\n')
                text_box.insert(END, 'Гидраты на забое образуются при температуре ниже ' + str(round(T,2)) + ' C')
        else:
            if 'Гидратообразование на забое скважины' in ogranicheniya:
                text_box.insert(END, "В скважине " + str(i) + " возможно образование гидратов\n")
                text_box.insert(END, 'Гидраты на забое образуются при температуре ниже ' + str(round(T,2)) + ' C')


    # ограничение по гидратам в шлейфе
    def hydrates_shleif(x, y):
        massa = []
        obem = []
        for i in range(x, y):
            global sum_mass
            global sum_obem
            massa.append(float(sheet['W' + str(i)].value) * int(sheet['F' + str(i)].value))
            sum_mass = np.sum(massa)
            obem.append(sheet['F' + str(i)].value)
            sum_obem = np.sum(obem)
        density = sum_mass / sum_obem
        e = density
        if sheet['N' + str(i)].value > 0:
            b = -12420 * e ** 5 + 50909 * e ** 4 - 83004 * e ** 3 + 67296 * e ** 2 - 27145 * e + 4375.1
            T = 18.47 * math.log10(sheet['M' + str(i)].value) - b + 18.47
        else:
            b1 = -39701 * e ** 5 + 162993 * e ** 4 - 266162 * e ** 3 + 216107 * e ** 2 - 87295 * e + 14090
            T = -58.51 * math.log10(sheet['M' + str(i)].value) - b1 - 58.51
        if T < sheet['N' + str(i)].value:
            if 'Гидратообразование в шлейфе' in ogranicheniya:
                text_box.insert(END, 'Текущая температура на входе в УКПГ = ' + str(round(sheet['N' + str(i)].value,2)) + ' C\n')
                text_box.insert(END, 'Гидраты в шлейфе скважин ' + str(x-3) + '-' + str(int(y-3) - 1) + ' образуются при температуре ниже ' + str(round(T,2)) + ' C\n')
        else:
            if 'Гидратообразование в шлейфе' in ogranicheniya:
                text_box.insert(END, "В шлейфе скважин " + str(x-3) + '-' + str(y-3) + " возможно образование гидратов\n")
                text_box.insert(END, 'Гидраты в шлейфе скважин ' + str(x-3) + '-' + str(int(y-3) - 1) + ' образуются при температуре ниже ' + str(round(T,2)) + ' C\n')

    def hydrates_shleif_muted(x, y):
        massa = []
        obem = []
        for i in range(x, y):
            global sum_mass
            global sum_obem
            massa.append(float(sheet['W' + str(i)].value) * int(sheet['F' + str(i)].value))
            sum_mass = np.sum(massa)
            obem.append(sheet['F' + str(i)].value)
            sum_obem = np.sum(obem)
        density = sum_mass / sum_obem
        e = density
        if sheet['N' + str(i)].value > 0:
            b = -12420 * e ** 5 + 50909 * e ** 4 - 83004 * e ** 3 + 67296 * e ** 2 - 27145 * e + 4375.1
            T = 18.47 * math.log10(sheet['M' + str(i)].value) - b + 18.47
        else:
            b1 = -39701 * e ** 5 + 162993 * e ** 4 - 266162 * e ** 3 + 216107 * e ** 2 - 87295 * e + 14090
            T = -58.51 * math.log10(sheet['M' + str(i)].value) - b1 - 58.51


    # print('Ограничение по коррозионным элементам:')
    # vehicle = []
    # vehicle_well(50)

    inputs_start = 4
    inputs_end = len(rasstoyanie_mejdu_skv) + 1
    Q_massiv = []
    pressure_potok_massiv = []
    pressure_potok_treb = []
    speed_potok_massiv = []

    # функция ограничения по давлению в шлейфе
    def pressure_shleif(x, y):
        vehicle = []
        pressure_shleif_min(inputs_start, inputs_start + inputs_end)
        hydrates_shleif_muted(inputs_start, inputs_start + inputs_end)
        pressure_vhod(inputs_start, inputs_start + inputs_end)
        for i in range(x, x+  1):  # цикл для первых двух скважин
            totalQ = 0
            ####расчет параметров для вычисления давления в начале i+1-ой скважины####
            global shleif_square
            global otnosit_sheroh
            global diameter
            shleif_square = 3.14 * ((sheet['Z' + str(i)].value) / 200) ** 2  # площадь сечения шлейфа
            otnosit_sheroh = sheet['AA' + str(i)].value / sheet[
                'Z' + str(i)].value  # относительная шероховатость шлейфа
            T_critical = 125 * (1 + sheet['W' + str(i)].value)
            P_critical = 0.4903 * (10 - sheet['W' + str(i)].value)
            T_priv = (sheet['K' + str(i)].value + 273.15) / T_critical
            P_priv = sheet['H' + str(i)].value / P_critical
            Z = (0.4 * math.log10(T_priv) + 0.73) ** P_priv + 0.1 * P_priv
            density = sheet['W' + str(i)].value * 1.2754
            density_priv = (sheet['W' + str(i)].value * 19.68 * sheet['J' + str(i)].value) / (
                        (sheet['K' + str(i)].value + 273.15) * Z)  # приведенная плотность
            densityPT = sheet['W' + str(i)].value * 3247.43 * sheet['J' + str(i)].value / (
                        Z * (sheet['K' + str(i)].value + 273.15))  # плотность при текущих P и Т
            dzeta = T_critical ** (1 / 6) / (4.583 * P_critical ** (2 / 3) * (
                        sheet['W' + str(i)].value * 29) ** 0.5)  # дзета для расчета вязкости
            if T_priv > 1.5:
                viscosity_initial = 166.8 * 0.00000001 * (0.1338 * T_priv - 0.0932) ** (
                            5 / 9) / dzeta  # вязкость при P0, T0,  Тприв>1.5
            else:
                viscosity_initial = 34 * 0.00000001 * T_priv ** (8 / 9) / dzeta  # вязкость при P0, T0,   Тприв<1.5
            viscosity_final = viscosity_initial + 10.8 * 0.00000001 * (math.exp(1.439 * density_priv) - math.exp(
                (-1.11) * density_priv ** 1.858)) / dzeta  # вязкость при текущих Р и Т
            diameter = sheet['Z' + str(i)].value  # диаметр шлейфа, см
            temperature = sheet['K' + str(i)].value + 273.15  # температура на входе в шлейфе
            pressure_skvajini_na_vhode_v_shleif = sheet['J' + str(i)].value * 1000000  # давление на выходе из скважины
            speed = sheet['F' + str(i)].value * 0.052 * temperature * Z / (
                        diameter ** 2 * sheet['J' + str(i)].value)  # скорость газа в шлейфе
            Re = speed * diameter / 100 * density / viscosity_final  # число рейнольдса
            gidravlich_trenie = 158 / Re + 2 * otnosit_sheroh  # коэффициент гидравлического трения лямбда
            gidravlich_trenie2 = 0.11 * (
                        otnosit_sheroh + 68 / Re) ** 0.25  # коэффициент гидравлического трения лямбда 2 версия
            R = 8314 / (sheet['W' + str(i)].value * 29)  # газовая постоянная для текущего состава газа
            M = 9.869 * (diameter / 100) ** 4 / 16 * densityPT ** 2 * speed ** 2  # массовый расход
            pressure_final = (pressure_skvajini_na_vhode_v_shleif ** 2 - 16 / 9.86960440108 * Z * R * temperature *
                              gidravlich_trenie2 * rasstoyanie_mejdu_skv[x - i] / (
                                          (diameter / 100) ** 5) * M ** 2) ** 0.5
            if isinstance(pressure_final, complex):
                pressure_final = 1000000
            print('Изменение давления от ' + str(i-3) + '-ой скважины до ' + str(i - 2) + '-ой:')
            print(str('От ' + str(sheet['J' + str(i)].value)) + ' до ' + str(pressure_final / 1000000) + ' МПа')
            print('Давление во входе в шлейф в ' + str(i -2) + '-ой скважине')
            print(sheet['J' + str(i+1)].value)
            ####конец расчета конечного давления от i-ой до i+1-ой скважины####
            ####начало расчета если давление больше, чем в i+1-ой скважине:####

            if pressure_final / 1000000 > sheet['J' + str(i + 1)].value:
                delta_P = pressure_final / 1000000 - sheet['J' + str(
                    i + 1)].value  # на сколько нужно изменить давление в начале шлейфа для выполнения условия оптимальности
                delta_Q = sheet['F' + str(i)].value * delta_P / sheet[
                    'J' + str(i)].value  # на сколько нужно изменить дебит для изменения давления
                new_Q = sheet['F' + str(i)].value - delta_Q  # новый дебит
                Q_massiv.append(new_Q)
                totalQ = np.sum(Q_massiv)
                ######преобразование переменных в формуле######
                newPressureZab = math.sqrt(sheet['E' + str(i)].value ** 2 - (sheet['E' + str(i)].value ** 2 - (
                        sheet['E' + str(i)].value - sheet['G' + str(i)].value) ** 2) * new_Q / sheet[
                                               'F' + str(i)].value)  # по новому дебиту находим новое забойное давление
                delta_P_nkt = newPressureZab - (
                        sheet['E' + str(i)].value - sheet['G' + str(i)].value)  # изменение давления в нкт и скважине
                P_priv_new = (sheet['H' + str(i)].value + delta_P_nkt) / P_critical
                Z_new = (0.4 * math.log10(T_priv) + 0.73) ** P_priv_new + 0.1 * P_priv_new
                density_priv_new = (sheet['W' + str(i)].value * 19.68 * (sheet['J' + str(i)].value - delta_P)) / (
                        (sheet['K' + str(i)].value + 273.15) * Z_new)  # приведенная плотность
                densityPT_new = sheet['W' + str(i)].value * 3247.43 * (sheet['J' + str(i)].value - delta_P) / (
                        Z_new * (sheet['K' + str(i)].value + 273.15))  # плотность при текущих P и Т
                viscosity_final_new = viscosity_initial + 10.8 * 0.00000001 * (
                            math.exp(1.439 * density_priv_new) - math.exp(
                        (-1.11) * density_priv_new ** 1.858)) / dzeta  # вязкость при текущих Р и Т
                pressure_skvajini_na_vhode_v_shleif_new = (sheet['J' + str(
                    i)].value - delta_P) * 1000000  # давление на выходе из скважины
                speed_new = new_Q * 0.052 * temperature * Z_new / (
                        diameter ** 2 * (sheet['J' + str(i)].value - delta_P))  # скорость газа в шлейфе
                Re = speed_new * diameter / 100 * density / viscosity_final_new  # число рейнольдса
                gidravlich_trenie = 158 / Re + 2 * otnosit_sheroh  # коэффициент гидравлического трения лямбда
                gidravlich_trenie2_new = 0.11 * (
                        otnosit_sheroh + 68 / Re) ** 0.25  # коэффициент гидравлического трения лямбда 2 версия
                M_new = 9.869 * (diameter / 100) ** 4 / 16 * densityPT_new ** 2 * speed_new ** 2  # массовый расход
                pressure_potok = ((
                                              pressure_skvajini_na_vhode_v_shleif_new ** 2 - 16 / 9.86960440108 * Z_new * R * temperature *
                                              gidravlich_trenie2_new * rasstoyanie_mejdu_skv[x - i] / (
                                                          (diameter / 100) ** 5) * M_new ** 2) ** 0.5) / 1000000
                pressure_potok_massiv.append(sheet['J' + str(i)].value - delta_P)
                pressure_potok_massiv.append(sheet['J' + str(i + 1)].value)
                pressure_potok_treb.append(sheet['J' + str(i)].value - delta_P)
                pressure_potok_treb.append(sheet['J' + str(i + 1)].value)

                ####конец расчета требуемого давления####
                ####начало расчета новых параметров потока####
                totalQ = totalQ + sheet[
                    'F' + str(i + 1)].value  # общий дебит в шлейфе 1 и 2 скважины после снижения давления
                global temperature_potok_C
                temperature_potok_C = new_Q / totalQ * sheet['K' + str(i)].value + sheet[
                    'F' + str(i + 1)].value / totalQ * sheet['K' + str(i + 1)].value
                temperature_potok_K = temperature_potok_C + 273.15
                density_potok_otn = new_Q / totalQ * sheet['W' + str(i)].value + sheet[
                    'F' + str(i + 1)].value / totalQ * sheet['W' + str(i + 1)].value
                density_potok = density_potok_otn * 1.2754
                T_critical_potok = 125 * (1 + density_potok_otn)
                P_critical_potok = 0.4903 * (10 - density_potok_otn)
                T_priv_potok = temperature_potok_K / T_critical_potok
                P_priv_potok = pressure_potok / P_critical_potok
                Z_potok = (0.4 * math.log10(T_priv_potok) + 0.73) ** P_priv_potok + 0.1 * P_priv_potok
                density_potok = density_potok_otn * 1.2754
                density_priv_potok = (density_potok_otn * 19.68 * pressure_potok) / (
                        temperature_potok_K * Z_potok)  # приведенная плотность
                densityPT_potok = density_potok_otn * 3247.43 * pressure_potok / (
                        Z_potok * temperature_potok_K)  # плотность при текущих P и Т
                dzeta_potok = T_critical_potok ** (1 / 6) / (4.583 * P_critical_potok ** (2 / 3) * (
                        density_potok_otn * 29) ** 0.5)  # дзета для расчета вязкости
                if T_priv > 1.5:
                    viscosity_initial_potok = 166.8 * 0.00000001 * (0.1338 * T_priv_potok - 0.0932) ** (
                            5 / 9) / dzeta_potok  # вязкость при P0, T0,  Тприв>1.5
                else:
                    viscosity_initial_potok = 34 * 0.00000001 * T_priv_potok ** (
                                8 / 9) / dzeta_potok  # вязкость при P0, T0,   Тприв<1.5
                viscosity_final_potok = viscosity_initial_potok + 10.8 * 0.00000001 * (
                            math.exp(1.439 * density_priv_potok) - math.exp(
                        (-1.11) * density_priv_potok ** 1.858)) / dzeta_potok  # вязкость при текущих Р и Т
                speed_potok = totalQ * 0.052 * temperature_potok_K * Z_potok / (
                        diameter ** 2 * pressure_potok)  # скорость газа в шлейфе
                speed_potok_massiv.append(speed_potok)
                Re_potok = speed * diameter / 100 * density_potok / viscosity_final_potok  # число рейнольдса
                gidravlich_trenie = 158 / Re + 2 * otnosit_sheroh  # коэффициент гидравлического трения лямбда
                gidravlich_trenie2 = 0.11 * (
                        otnosit_sheroh + 68 / Re_potok) ** 0.25  # коэффициент гидравлического трения лямбда 2 версия
                R_potok = 8314 / (density_potok_otn * 29)  # газовая постоянная для текущего состава газа
                M_potok = 3.14 * (diameter / 100) ** 2 / 4 * densityPT_potok * speed_potok  # массовый расход
                print('Новое давление после преобразования возле ' + str(i -2) + '-ой скважины : ' + str(
                    pressure_potok) + ' МПа')
                print('#############################')
                print(pressure_potok)
                print('Новая скорость в шлейфе после ' + str(i -2) + ' скважины: ' + str(speed_potok) + 'м/c')
                print('Новое забойное давление ' + str(newPressureZab) + ' МПа')
                print('Общий дебит ' + str(totalQ) + ' тыс.м3/сут')
                print('Температура потока: ' + str(temperature_potok_K - 273.15) + ' С')
                for element in pressure_potok_massiv:
                    print(str(element))
                Q_massiv.append(sheet['F' + str(i + 1)].value)
                ####начало расчета если давление меньше, чем в i+1-ой скважине:#### (тупо прибавляем)
            else:
                new_Q = sheet['F' + str(i)].value  # новый дебит
                Q_massiv.append(sheet['F' + str(i)].value)
                Q_massiv.append(sheet['F' + str(i+1)].value)
                totalQ = np.sum(Q_massiv)
                temperature_potok_C = sheet['F' + str(i)].value / totalQ * sheet['K' + str(i)].value + sheet[
                    'F' + str(i + 1)].value / totalQ * sheet['K' + str(i + 1)].value
                temperature_potok_K = temperature_potok_C + 273.15
                density_potok_otn = sheet['F' + str(i)].value / totalQ * sheet['W' + str(i)].value + sheet[
                    'F' + str(i + 1)].value / totalQ * sheet['W' + str(i + 1)].value
                density_potok = density_potok_otn * 1.2754
                T_critical_potok = 125 * (1 + density_potok_otn)
                P_critical_potok = 0.4903 * (10 - density_potok_otn)
                T_priv_potok = temperature_potok_K / T_critical_potok
                P_priv_potok = sheet['J' + str(i)].value / P_critical_potok
                Z_potok = (0.4 * math.log10(T_priv_potok) + 0.73) ** P_priv_potok + 0.1 * P_priv_potok
                density_potok = density_potok_otn * 1.2754
                density_priv_potok = (density_potok_otn * 19.68 * sheet['J' + str(i)].value) / (
                        temperature_potok_K * Z_potok)  # приведенная плотность
                densityPT_potok = density_potok_otn * 3247.43 * sheet['J' + str(i)].value / (
                        Z_potok * temperature_potok_K)  # плотность при текущих P и Т
                dzeta_potok = T_critical_potok ** (1 / 6) / (4.583 * P_critical_potok ** (2 / 3) * (
                        density_potok_otn * 29) ** 0.5)  # дзета для расчета вязкости
                if T_priv > 1.5:
                    viscosity_initial_potok = 166.8 * 0.00000001 * (0.1338 * T_priv_potok - 0.0932) ** (
                            5 / 9) / dzeta_potok  # вязкость при P0, T0,  Тприв>1.5
                else:
                    viscosity_initial_potok = 34 * 0.00000001 * T_priv_potok ** (
                                8 / 9) / dzeta_potok  # вязкость при P0, T0,   Тприв<1.5
                viscosity_final_potok = viscosity_initial_potok + 10.8 * 0.00000001 * (
                            math.exp(1.439 * density_priv_potok) - math.exp(
                        (-1.11) * density_priv_potok ** 1.858)) / dzeta_potok  # вязкость при текущих Р и Т
                speed_potok = totalQ * 0.052 * temperature_potok_K * Z_potok / (
                        diameter ** 2 * sheet['J' + str(i)].value)  # скорость газа в шлейфе
                speed_potok_massiv.append(speed_potok)
                Re_potok = speed * diameter / 100 * density_potok / viscosity_final_potok  # число рейнольдса
                gidravlich_trenie = 158 / Re + 2 * otnosit_sheroh  # коэффициент гидравлического трения лямбда
                gidravlich_trenie2 = 0.11 * (
                        otnosit_sheroh + 68 / Re_potok) ** 0.25  # коэффициент гидравлического трения лямбда 2 версия
                R_potok = 8314 / (density_potok_otn * 29)  # газовая постоянная для текущего состава газа
                M_potok = 3.14 * (diameter / 100) ** 2 / 4 * densityPT_potok * speed_potok  # массовый расход
                pressure_potok = sheet['J' + str(i)].value
                pressure_potok_massiv.append(pressure_potok)
                pressure_potok_massiv.append(sheet['J' + str(i + 1)].value)
                pressure_potok_treb.append(pressure_potok)
                pressure_potok_treb.append(sheet['J' + str(i + 1)].value)
                print('4444444444444444444444444444444')
                print('новое давление после преобразования:' + str(pressure_potok) + ' МПа')
                print('новая скорость в шлейфе после ' + str(i -2) + ' скважины: ' + str(speed_potok) + 'м/c')
                print('забойное давление не изменится')
                print('общий дебит: ' + str(totalQ) + ' тыс.м3/сут')
                print('температура потока: ' + str(temperature_potok_C) + ' C')
                for element in pressure_potok_massiv:
                    print(str(element))
            for i in range(x + 1, y - 1):  # расчет потока для 2+ скважины
                print('uuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuu')
                print(pressure_potok_massiv[i - x])
                print(Z_potok)
                print(R_potok)
                print(temperature_potok_K)
                print(gidravlich_trenie2)
                print(rasstoyanie_mejdu_skv[x + 1 - i])
                print(diameter)
                print(M_potok)
                if ((pressure_potok_massiv[i - x] * 1000000) ** 2) < (16 / 9.86960440108 * Z_potok * R_potok * temperature_potok_K * gidravlich_trenie2 * rasstoyanie_mejdu_skv[x + 1 - i] / ((diameter / 100) ** 5) * M_potok ** 2):
                    pressure_potok_new = 1
                else:
                    pressure_potok_new = (((pressure_potok_massiv[i - x] * 1000000) ** 2 - 16 / 9.86960440108 * Z_potok * R_potok * temperature_potok_K * gidravlich_trenie2 * rasstoyanie_mejdu_skv[x + 1 - i] / ((diameter / 100) ** 5) * M_potok ** 2) ** 0.5) / 1000000
                if isinstance(pressure_potok_new, complex):
                    pressure_potok_new = 1
                if pressure_potok_new < sheet[
                    'J' + str(i + 1)].value:  # если P в скважине > P в шлейфе(тупо прибавляем дебит)
                    Q_massiv.append(sheet['F' + str(i + 1)].value)
                    totalQ = np.sum(Q_massiv)
                    temperature_potok_C = (totalQ - sheet['F' + str(i + 1)].value) / totalQ * temperature_potok_C + \
                                          sheet['F' + str(i + 1)].value / totalQ * sheet['K' + str(i + 1)].value
                    temperature_potok_K = temperature_potok_C + 273.15
                    density_potok_otn = (totalQ - sheet['F' + str(i)].value) / totalQ * density_potok_otn + sheet[
                        'F' + str(i + 1)].value / totalQ * sheet['W' + str(i + 1)].value
                    density_potok = density_potok_otn * 1.2754
                    T_critical_potok = 125 * (1 + density_potok_otn)
                    P_critical_potok = 0.4903 * (10 - density_potok_otn)
                    T_priv_potok = temperature_potok_K / T_critical_potok
                    P_priv_potok = sheet['J' + str(i + 1)].value / P_critical_potok
                    Z_potok = (0.4 * math.log10(T_priv_potok) + 0.73) ** P_priv_potok + 0.1 * P_priv_potok
                    density_priv_potok = (density_potok_otn * 19.68 * sheet['J' + str(i + 1)].value) / (
                            temperature_potok_K * Z_potok)  # приведенная плотность
                    densityPT_potok = density_potok_otn * 3247.43 * sheet['J' + str(i + 1)].value / (
                            Z_potok * temperature_potok_K)  # плотность при текущих P и Т
                    dzeta_potok = T_critical_potok ** (1 / 6) / (4.583 * P_critical_potok ** (2 / 3) * (
                            density_potok_otn * 29) ** 0.5)  # дзета для расчета вязкости
                    if T_priv > 1.5:
                        viscosity_initial_potok = 166.8 * 0.00000001 * (0.1338 * T_priv_potok - 0.0932) ** (
                                5 / 9) / dzeta_potok  # вязкость при P0, T0,  Тприв>1.5
                    else:
                        viscosity_initial_potok = 34 * 0.00000001 * T_priv_potok ** (
                                8 / 9) / dzeta_potok  # вязкость при P0, T0,   Тприв<1.5
                    viscosity_final_potok = viscosity_initial_potok + 10.8 * 0.00000001 * (
                            math.exp(1.439 * density_priv_potok) - math.exp(
                        (-1.11) * density_priv_potok ** 1.858)) / dzeta_potok  # вязкость при текущих Р и Т
                    speed_potok = totalQ * 0.052 * temperature_potok_K * Z_potok / (
                            diameter ** 2 * sheet['J' + str(i + 1)].value)  # скорость газа в шлейфе
                    speed_potok_massiv.append(speed_potok)
                    Re_potok = speed_potok * diameter / 100 * density_potok / viscosity_final_potok  # число рейнольдса
                    gidravlich_trenie = 158 / Re_potok + 2 * otnosit_sheroh  # коэффициент гидравлического трения лямбда
                    gidravlich_trenie2 = 0.11 * (
                            otnosit_sheroh + 68 / Re_potok) ** 0.25  # коэффициент гидравлического трения лямбда 2 версия
                    R_potok = 8314 / (density_potok_otn * 29)  # газовая постоянная для текущего состава газа
                    M_potok = 3.14 * (diameter / 100) ** 2 / 4 * densityPT_potok * speed_potok  # массовый расход
                    pressure_potok = sheet['J' + str(i + 1)].value
                    pressure_potok_massiv.append(pressure_potok)
                    pressure_potok_treb.append(pressure_potok_new)
                    print('%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%')
                    print('новое давление после преобразования:' + str(pressure_potok) + ' МПа')
                    print('новая скорость в шлейфе после ' + str(i -2) + ' скважины: ' + str(speed_potok) + 'м/c')
                    print('забойное давление скважины' + str(i -2) + 'не изменится')
                    print(
                        'температура потока после ' + str(i -2) + ' скважины: ' + str(temperature_potok_C) + ' C')
                    print('дебит после скважины ' + str(i -2) + ': ' + str(sum(Q_massiv)) + ' тыс. м3/сут')
                    for element in Q_massiv:
                        print(str(element) + ' тыс.м3/сут')
                else:  # pressure_potok_new > sheet['J' + str(i+1)].value(давление в шлейфе больше, чем давление в скв => снижаем давление во всех скв.)
                    diff = sheet['J' + str(i + 1)].value - pressure_potok_new
                    pressure_potok_massiv.append(pressure_potok)
                    Q_massiv.append(sheet['F' + str(i + 1)].value)
                    for l in range(0, len(pressure_potok_massiv)):
                        pressure_potok_massiv[l] = pressure_potok_massiv[l] + diff
                    for l in range(0, len(pressure_potok_massiv)):
                        Q_massiv[l] = Q_massiv[l] * pressure_potok_massiv[l] / (pressure_potok_massiv[l] - diff)

                    temperature_potok_C = (sum(Q_massiv) - sheet['F' + str(i + 1)].value) / sum(
                        Q_massiv) * temperature_potok_C + sheet[
                                              'F' + str(i + 1)].value / sum(Q_massiv) * sheet['K' + str(i + 1)].value
                    temperature_potok_K = temperature_potok_C + 273.15
                    density_potok_otn = (sum(Q_massiv) - sheet['F' + str(i)].value) / sum(
                        Q_massiv) * density_potok_otn + sheet[
                                            'F' + str(i + 1)].value / sum(Q_massiv) * sheet['W' + str(i + 1)].value
                    density_potok = density_potok_otn * 1.2754
                    T_critical_potok = 125 * (1 + density_potok_otn)
                    P_critical_potok = 0.4903 * (10 - density_potok_otn)
                    T_priv_potok = temperature_potok_K / T_critical_potok
                    P_priv_potok = sheet['J' + str(i + 1)].value / P_critical_potok
                    Z_potok = (0.4 * math.log10(T_priv_potok) + 0.73) ** P_priv_potok + 0.1 * P_priv_potok
                    density_priv_potok = (density_potok_otn * 19.68 * sheet['J' + str(i + 1)].value) / (
                            temperature_potok_K * Z_potok)  # приведенная плотность
                    densityPT_potok = density_potok_otn * 3247.43 * sheet['J' + str(i + 1)].value / (
                            Z_potok * temperature_potok_K)  # плотность при текущих P и Т
                    dzeta_potok = T_critical_potok ** (1 / 6) / (4.583 * P_critical_potok ** (2 / 3) * (
                            density_potok_otn * 29) ** 0.5)  # дзета для расчета вязкости
                    if T_priv > 1.5:
                        viscosity_initial_potok = 166.8 * 0.00000001 * (0.1338 * T_priv_potok - 0.0932) ** (
                                5 / 9) / dzeta_potok  # вязкость при P0, T0,  Тприв>1.5
                    else:
                        viscosity_initial_potok = 34 * 0.00000001 * T_priv_potok ** (
                                8 / 9) / dzeta_potok  # вязкость при P0, T0,   Тприв<1.5
                    viscosity_final_potok = viscosity_initial_potok + 10.8 * 0.00000001 * (
                            math.exp(1.439 * density_priv_potok) - math.exp(
                        (-1.11) * density_priv_potok ** 1.858)) / dzeta_potok  # вязкость при текущих Р и Т
                    speed_potok = sum(Q_massiv) * 0.052 * temperature_potok_K * Z_potok / (
                            diameter ** 2 * sheet['J' + str(i + 1)].value)  # скорость газа в шлейфе
                    speed_potok_massiv.append(speed_potok)
                    Re_potok = speed_potok * diameter / 100 * density_potok / viscosity_final_potok  # число рейнольдса
                    gidravlich_trenie = 158 / Re_potok + 2 * otnosit_sheroh  # коэффициент гидравлического трения лямбда
                    gidravlich_trenie2 = 0.11 * (
                            otnosit_sheroh + 68 / Re_potok) ** 0.25  # коэффициент гидравлического трения лямбда 2 версия
                    R_potok = 8314 / (density_potok_otn * 29)  # газовая постоянная для текущего состава газа
                    M_potok = 3.14 * (diameter / 100) ** 2 / 4 * densityPT_potok * speed_potok  # массовый расход
                    pressure_potok = sheet['J' + str(i + 1)].value
#                    pressure_potok_massiv.append(pressure_potok)
                    pressure_potok_treb.append(pressure_potok_new)
                    print('66666666666666666666666666666666666666666666')
                    print(sum(Q_massiv))
                    print(temperature_potok_K)
                    print(Z_potok)
                    print(diameter)

                    print('новое давление после преобразования:' + str(pressure_potok) + ' МПа')
                    print('новая скорость в шлейфе после ' + str(i -2) + ' скважины: ' + str(speed_potok) + 'м/c')
                    print('забойное давление скважины' + str(i -2) + 'не изменится')
                    print(
                        'температура потока после ' + str(i -2) + ' скважины: ' + str(temperature_potok_C) + ' C')
                    print('дебит после скважины ' + str(i -2) + ': ' + str(sum(Q_massiv)) + ' тыс. м3/сут')
                    for element in Q_massiv:
                        print(str(element) + ' тыс.м3/сут')
                    pressure_potok_treb.append(pressure_potok_new)
#                    pressure_potok_massiv.append(
#                        pressure_potok)  # pressure_potok_treb должен быть < pressure_potok_massive
                    for element in Q_massiv:
                        print(str(element) + ' тыс.м3/сут')


    Q_massiv_speed = []

    def speed_potok_func(x):  # ограничение по скорости потока в шлейфе 17 мс
        if x > 17:
            for l in range(0, len(rasstoyanie_mejdu_skv)+1):
                Q_massiv[l] = 17 / speed_potok_massiv[-1] * Q_massiv[l]
                Q_massiv_speed.append(Q_massiv[l])
            print('Требуется ограничение по скорости в шлейфе ')
            for element in Q_massiv_speed:
                print(str(element) + ' тыс.м3/сут')
                print('\n')
            print('\n')
        else:
            for l in range(0, len(Q_massiv)):
                print('Текущая скорость потока: ' + str(speed_potok_massiv[-1]) + 'м/c')
                Q_massiv_speed.append(Q_massiv[l])
                for element in Q_massiv_speed:
                    print(str(element))
                    print('\n')
                print('Ограничение по скорости в шлейфе не требуется')
                print('\n')
        if 'Ограничение скорости потока в шлейфе' in ogranicheniya:
            if speed_potok_massiv[-1] > 17:
                text_box.insert(END, '\n\nМаксимальная скорость потока  = 17 м/с\n')
                text_box.insert(END, 'Текущая скорость потока в шлейфе = ' + str(speed_potok_massiv[-1]) + ' м/с\n')
                text_box.insert(END, 'Требуется ограничение по скорости в шлейфе\n')
                for i in range(len(rasstoyanie_mejdu_skv)+1):
                    text_box.insert(END, 'Ограничения по дебиту в скважине ' + str(i) + ' с новой скоростью в шлейфе = ' + str(round(Q_massiv_speed[i],2)) + ' тыс.м3/сут\n')
            elif speed_potok_massiv[-1] < 17:
                text_box.insert(END, '\n\nМаксимальная скорость потока  = 17 м/с\n')
                text_box.insert(END, 'Текущая скорость потока в шлейфе = ' + str(speed_potok_massiv[-1]) + ' м/с\n')
                text_box.insert(END, 'Ограничение по скорости в шлейфе не требуется\n')


    Q_dopustimiy_rejim = []

    def dopustimiy_rejim(x):  # ограничение по дебиту в допустимом режиме
        summa = np.sum(x)
        avg = summa / 6
        chast = avg / 5
        for l in range(0, len(x)):
            Q_dopustimiy_rejim.append(x[l] + chast)
        print('Ограничение по допустимому режиму:')
        for element in Q_dopustimiy_rejim:
            print(str(element) + ' тыс.м3/сут')



    # расчет давления при новом оптимальном режиме
    P_zab1_opt_massiv = []
    P_zab2_opt_massiv = []

    def P_zab_opt(x, y):
        for i in range(x, y):
            Pzab2 = math.sqrt(sheet['E' + str(i)].value ** 2 - (
                        sheet['E' + str(i)].value ** 2 - (sheet['E' + str(i)].value - sheet['G' + str(i)].value) ** 2) *
                              Q_massiv_speed[i - x] / sheet['F' + str(i)].value)
            P_zab2_opt_massiv.append(Pzab2)
            P_zab1_opt_massiv.append(sheet['E' + str(i)].value - sheet['G' + str(i)].value)
        print('Давления на забое до ограничений по скорости в шлейфе:')
        print('\n')
        for element in P_zab1_opt_massiv:
            print(str(element) + ' МПа')
            print('\n')
        print('Давления на забое после ограничений по скорости в шлейфе:')
        print('\n')
        for element in P_zab2_opt_massiv:
            print(str(element) + ' МПа')
            print('\n')

    # расчет давления при новом допустимом режиме
    P_zab2_dop_massiv = []

    def P_zab_dop(x, y):
        for i in range(x, y):
            Pzab2 = math.sqrt(sheet['E' + str(i)].value ** 2 - (
                        sheet['E' + str(i)].value ** 2 - (sheet['E' + str(i)].value - sheet['G' + str(i)].value) ** 2) *
                              Q_dopustimiy_rejim[i - x] / sheet['F' + str(i)].value)
            P_zab2_dop_massiv.append(Pzab2)
        print('Забойные давления при допустимом режиме:')
        for element in P_zab2_dop_massiv:
            print(str(element) + ' МПа')



    # if 'Максимальный безводный дебит' in ogranicheniya:

    # if 'Гидратообразование в стволе скважины' in ogranicheniya:

    pressure_shleif(inputs_start, inputs_start + inputs_end)
    speed_potok_func(speed_potok_massiv[-1])
    dopustimiy_rejim(Q_massiv_speed)
    P_zab_dop(inputs_start, inputs_start+inputs_end)
    P_zab_opt(inputs_start, inputs_start+inputs_end)
    hydrates_shleif(inputs_start, inputs_start+inputs_end)

    def relu(x):
        return np.maximum(0, x)

    print(whites)
    print(reds)
    np.random.seed(1)
    synaptic_weights = 2 * np.random.random((12, 1))
    text_box.insert(END, '\n\nСлучайные веса:\n')
    for i in range(0, len(synaptic_weights)):
        text_box.insert(END, synaptic_weights[i])
        text_box.insert(END, '\n')

    for i in range(0, 11):
        if i in reds:
            synaptic_weights[i - 1] = 0
    if 11 in reds:
        synaptic_weights[-2] = 0
    if 12 in reds:
        synaptic_weights[-1] = 0

    training_inputs_noDATA = training_inputs2
    for i in range(0, len(training_inputs2)):
        element = training_inputs_noDATA[i][0]
        training_inputs_noDATA[i].remove(element)

    text_box.insert(END, '\n')
    #text_box.insert(END, time.clock())
    time1 = time.clock()
    # вычисление весов для оптимального режима работы скважины
    training_outputs_optimal2 = [sheet['R4'].value]
    for j in range(5, 50):
        training_outputs_optimal2.append(sheet['R' + str(j)].value)
    training_outputs_optimal = np.array(training_outputs_optimal2).T
    synaptic_weights_optimal = synaptic_weights
    training_inputs = np.array(training_inputs_noDATA)
    # обучение
    iterations = Entry17.get()
    for k in range(int(iterations)):
        input_layer = training_inputs
        outputs_optimal = relu(np.dot(input_layer, synaptic_weights_optimal))
        err = training_outputs_optimal - outputs_optimal
        for i in range(0, 11):
            if synaptic_weights_optimal[i] == 0:
                synaptic_weights_optimal[i] = 0
            else:
                synaptic_weights_optimal[i] = synaptic_weights_optimal[i] + np.mean(
                    err * (1 / (12 - len(reds))) / training_outputs_optimal)

    text_box.insert(END, '\nВеса после обучения (' + str(iterations) + ' итераций) для оптимального режима работы:\n')
    text_box.insert(END, synaptic_weights_optimal)

    print(whites)
    print(reds)
    np.random.seed(1)
    synaptic_weights = 1.9 * np.random.random((12, 1))
    text_box.insert(END, '\n\nСлучайные веса:\n')
    for i in range(0, len(synaptic_weights)):
        text_box.insert(END, synaptic_weights[i])
        text_box.insert(END, '\n')

    for i in range(0, 11):
        if i in reds:
            synaptic_weights[i - 1] = 0
    if 11 in reds:
        synaptic_weights[-2] = 0
    if 12 in reds:
        synaptic_weights[-1] = 0

    # вычисление весов для допустимого режима работы скважины
    synaptic_weights_dopustim = synaptic_weights
    training_outputs_dopustim2 = [sheet['U4'].value]
    for j in range(5, 50):
        training_outputs_dopustim2.append(sheet['U' + str(j)].value)
    training_outputs_dopustim = np.array(training_outputs_dopustim2).T
    for k in range(int(iterations)):
        input_layer = training_inputs
        outputs_dopustim = relu(np.dot(input_layer, synaptic_weights_dopustim))
        err = training_outputs_dopustim - outputs_dopustim
        for i in range(0, 11):
            if synaptic_weights_dopustim[i] == 0:
                synaptic_weights_dopustim[i] = 0
            else:
                synaptic_weights_dopustim[i] = synaptic_weights_dopustim[i] + np.mean(
                    err * (1 / (12 - len(reds))) / training_outputs_dopustim)
    text_box.insert(END, '\n')
    #text_box.insert(END, time.clock())
    time2 = time.clock()

    text_box.insert(END, '\n\nВеса после обучения (' + str(iterations) + ' итераций) для допустимого режима работы:\n')
    text_box.insert(END, synaptic_weights_dopustim)
    # print('pppppppppppppppppp')
    # print(Q_dopustimiy_rejim)
    # print(speed_potok_massiv)
    # print(P_zab2_dop_massiv)
    text_box.insert(END, '\nВремя обучения нейронной сети:\n')
    text_box.insert(END, str(time2-time1) + 'сек.\n')
    for i in range(inputs_start, inputs_start + inputs_end):
        text_box.insert(END, '\n\nСКВАЖИНА №' + str(i-3) + '\n')
        if i != (inputs_start + inputs_end-1):
            density = sum_mass / sum_obem
            e = density
            if sheet['N' + str(i)].value > 0:
                b = -12420 * e ** 5 + 50909 * e ** 4 - 83004 * e ** 3 + 67296 * e ** 2 - 27145 * e + 4375.1
                T = 18.47 * math.log10(sheet['M' + str(i)].value) - b + 18.47
            else:
                b1 = -39701 * e ** 5 + 162993 * e ** 4 - 266162 * e ** 3 + 216107 * e ** 2 - 87295 * e + 14090
                T = -58.51 * math.log10(sheet['M' + str(i)].value) - b1 - 58.51
            if T < sheet['N' + str(i)].value:
                if 'Гидратообразование в шлейфе' in ogranicheniya:
                    text_box.insert(END,
                                    '\nТекущая температура на входе в УКПГ = ' + str(round(
                                        sheet['N' + str(i)].value,2)) + ' C\n')
                    text_box.insert(END, 'Гидраты в шлейфе скважин ' + str(i - 3) + '-' + str(
                        int(i) - 2) + ' образуются при температуре ниже ' + str(round(T,2)) + ' C')
            else:
                if 'Гидратообразование в шлейфе' in ogranicheniya:
                    text_box.insert(END,
                                    "В шлейфе скважин " + str(i - 3) + '-' + str(
                                        i - 2) + " возможно образование гидратов\n")
                    text_box.insert(END, 'Гидраты в шлейфе скважин ' + str(i - 3) + '-' + str(
                        int(i - 1) - 1) + ' образуются при температуре ниже ' + str(round(T,2)) + ' C\n')
        else:
            pass

        new_inputs = np.array([sheet['D' + str(i)].value, sheet['E' + str(i)].value, sheet['F' + str(i)].value,
                               sheet['G' + str(i)].value, sheet['H' + str(i)].value, sheet['I' + str(i)].value,
                               pipe_pressure, sheet['K' + str(i)].value, sheet['L' + str(i)].value,
                               sheet['M' + str(i)].value, sheet['N' + str(i)].value, sheet['O' + str(i)].value])
#        text_box.insert(END, synaptic_weights_optimal)
#        text_box.insert(END, '\n')
#        text_box.insert(END, synaptic_weights_dopustim)
        new_output_optimal = relu(np.dot(new_inputs, synaptic_weights_optimal))
        text_box.insert(END, '\nОптимальный дебит скважины по нейросети =  ' + str("%.3f" % new_output_optimal) + ' тыс.м3/сут\n')
        new_output_dopustim = relu(np.dot(new_inputs, synaptic_weights_dopustim))
        text_box.insert(END, 'Допустимый дебит скважины по нейросети: ' + str("%.3f" % new_output_dopustim) + ' тыс.м3/сут')
        # ИНДИВИДУАЛЬНЫЕ ОГРАНИЧЕНИЯ
        hydrates_wellhead(i)
        hydrates_bottom(i)
        print('######################################################')


        water_probka(i)
        obvodnenie(i)
        condensat_probka(i)
        vehicle_well(i)

        zaboynoe_davlenie = math.sqrt(sheet['E' + str(i)].value ** 2 - (sheet['E' + str(i)].value ** 2 - (
                sheet['E' + str(i)].value - sheet['G' + str(i)].value) ** 2) * new_output_optimal / sheet[
                                          'F' + str(i)].value)

        def predict_P():
            mesyaci = int(Entry18.get())
            i = inputs_end
            # Накопленный дебит по макс. доп. режиму:
            Qnak_dop = []
            for k in range(0, mesyaci):

                Qnak_dop_mes = []
#тут должно быть k*6, потом расширить таблицу, т.к. не хватает данных
                for j in range(18+k, 18+i+k):
                #Qдоп в первой скважине 1 месяца
                    sheet['U' + str(j)].value = sheet['U' + str(j - 6)].value * (sheet['U' + str(j - 6)].value
                                                                             + sheet['U' + str(j - 5)].value + sheet[
                                                                                 'U' + str(j - 4)].value + sheet[
                                                                                 'U' + str(j - 3)].value
                                                                             + sheet['U' + str(j - 2)].value + sheet[
                                                                                 'U' + str(j - 1)].value) / \
                                            (sheet['U' + str(j - 7)].value + sheet['U' + str(j - 8)].value + sheet[
                                                'U' + str(j - 9)].value
                                             + sheet['U' + str(j - 10)].value + sheet['U' + str(j - 11)].value + sheet[
                                                 'U' + str(j - 12)].value)
                    Qnak_dop_mes.append(sheet['U' + str(j)].value)
                if ('Дебит накопленный по максимально допустимому режиму на ' + Entry18.get() + ' месяцев') in prognozi:
                    text_box.insert(END, '\nДебит, накопленный по максимально допустимому режиму работы за ' + str(k+1) + '-й месяц на скважинах равен ' + str(round(sum(Qnak_dop_mes),2)) + ' тыс.м3 газа')
                Qnak_dop.append(sum(Qnak_dop_mes))

            if ('Дебит накопленный по максимально допустимому режиму на ' + Entry18.get() + ' месяцев') in prognozi:
                text_box.insert(END, '\nСуммарный накопленный дебит по максимально допустимому режиму  работы на выбранных скважинах за ' + str(Entry18.get()) + ' месяцев равен ' + str(round(sum(Qnak_dop),2)) + ' тыс. м3 газа\n')
        # Накопленный дебит по оптим. режиму:
            Qnak_dop = []
            for k in range(0, mesyaci):
                Qnak_dop_mes = []
            # тут должно быть k*6, потом расширить таблицу, т.к. не хватает данных
                for j in range(18 + k, 18 + i + k):
                # Qдоп в первой скважине 1 месяца
                    sheet['R' + str(j)].value = sheet['R' + str(j - 6)].value * (sheet['R' + str(j - 6)].value
                                                                             + sheet['R' + str(j - 5)].value + sheet[
                                                                                 'R' + str(j - 4)].value + sheet[
                                                                                 'R' + str(j - 3)].value
                                                                             + sheet['R' + str(j - 2)].value + sheet[
                                                                                 'R' + str(j - 1)].value) / \
                                            (sheet['R' + str(j - 7)].value + sheet['R' + str(j - 8)].value + sheet[
                                                'R' + str(j - 9)].value
                                             + sheet['R' + str(j - 10)].value + sheet['R' + str(j - 11)].value + sheet[
                                                 'R' + str(j - 12)].value)
                    Qnak_dop_mes.append(sheet['R' + str(j)].value)
                if ('Дебит накопленный по оптимальному режиму на ' + Entry18.get() + ' месяцев') in prognozi:
                    text_box.insert(END, '\nДебит, накопленный по оптимальному режиму работы за ' + str(
                    k + 1) + '-й месяц на скважинах равен ' + str(round(sum(Qnak_dop_mes),2)) + ' тыс.м3 газа')
                    Qnak_dop.append(sum(Qnak_dop_mes))
            if ('Дебит накопленный по оптимальному режиму на ' + Entry18.get() + ' месяцев') in prognozi:
                text_box.insert(END, '\nСуммарный накопленный дебит по оптимальному режиму работы на выбранных скважинах за ' + str(
                    Entry18.get()) + ' месяцев равен ' + str(round(sum(Qnak_dop),2)) + ' тыс. м3 газа\n')

            minimals = []
            minimals.append(Q_massiv_speed[i - 4])
            minimals.append(q_water_min_minimals[i - 4])
            minimals.append(Q_bezvodn_minimals[i - 4])
            minimals.append(q_condensat_min_minimals[i - 4])
            minimals.append(q_critical_5_minimals[i - 4])
            text_box.insert(END, '\n\nДебит по минимальному допустимому режиму на конец ' + str(k + 1) + '-го месяца = ' + str(
                            round(min(minimals), 2)) + ' тыс.м3/сут\n\n')
            Qnak_min = sheet['R' + str(j)].value



            #депрессия при доп режиме
            for k in range(0, mesyaci):
                for j in range(57 + k, 57 + i + k):
                    sheet['T' + str(j)].value = sheet['T' + str(j - 6)].value * (sheet['T' + str(j - 6)].value
                                                                                     + sheet['T' + str(j - 5)].value +
                                                                                     sheet[
                                                                                         'T' + str(j - 4)].value +
                                                                                     sheet[
                                                                                         'T' + str(j - 3)].value
                                                                                     + sheet['T' + str(j - 2)].value +
                                                                                     sheet[
                                                                                         'T' + str(j - 1)].value) / \
                                                    (sheet['T' + str(j - 7)].value + sheet['T' + str(j - 8)].value +
                                                     sheet[
                                                         'T' + str(j - 9)].value
                                                     + sheet['T' + str(j - 10)].value + sheet['T' + str(j - 11)].value +
                                                     sheet[
                                                         'T' + str(j - 12)].value)
            # депрессия при опт режиме
            for k in range(0, mesyaci):
                for j in range(57 + k, 57 + i + k):
                    sheet['Q' + str(j)].value = sheet['Q' + str(j - 6)].value * (sheet['Q' + str(j - 6)].value
                                                                                         + sheet[
                                                                                             'Q' + str(j - 5)].value +
                                                                                         sheet[
                                                                                             'Q' + str(j - 4)].value +
                                                                                         sheet[
                                                                                             'Q' + str(j - 3)].value
                                                                                         + sheet[
                                                                                             'Q' + str(j - 2)].value +
                                                                                         sheet[
                                                                                             'Q' + str(j - 1)].value) / \
                                                        (sheet['Q' + str(j - 7)].value + sheet['Q' + str(j - 8)].value +
                                                         sheet[
                                                             'Q' + str(j - 9)].value
                                                         + sheet['Q' + str(j - 10)].value + sheet[
                                                             'Q' + str(j - 11)].value +
                                                         sheet[
                                                             'Q' + str(j - 12)].value)
                #давление устьевое при ОПТИМ режиме
            for k in range(0, mesyaci):
                for j in range(57 + k, 57 + i + k):
                    sheet['P' + str(j)].value = sheet['P' + str(j - 6)].value * (sheet['P' + str(j - 6)].value
                                                                                     + sheet['P' + str(j - 5)].value +
                                                                                     sheet[
                                                                                         'P' + str(j - 4)].value +
                                                                                     sheet[
                                                                                         'P' + str(j - 3)].value
                                                                                     + sheet['P' + str(j - 2)].value +
                                                                                     sheet[
                                                                                         'P' + str(j - 1)].value) / \
                                                    (sheet['P' + str(j - 7)].value + sheet['P' + str(j - 8)].value +
                                                     sheet[
                                                         'P' + str(j - 9)].value
                                                     + sheet['P' + str(j - 10)].value + sheet['P' + str(j - 11)].value +
                                                     sheet[
                                                         'P' + str(j - 12)].value)
                    if 'Изменение устьевого давления Pу при оптимальном режиме на ' + Entry18.get() + ' месяцев' in prognozi:
                        text_box.insert(END, 'Устьевое давление скважины ' + str(j-57) + ' при оптимальном режиме на конец ' + str(k+1) + '-го месяца = ' + str(sheet['P' + str(j)].value) + ' МПа\n')
                # давление устьевое при макс доп режиме
            for k in range(0, mesyaci):
                for j in range(57 + k, 57 + i + k):
                    sheet['S' + str(j)].value = sheet['S' + str(j - 6)].value * (
                                            sheet['S' + str(j - 6)].value
                                            + sheet['S' + str(j - 5)].value +
                                            sheet[
                                                'S' + str(j - 4)].value +
                                            sheet[
                                                'S' + str(j - 3)].value
                                            + sheet['S' + str(j - 2)].value +
                                            sheet[
                                                'S' + str(j - 1)].value) / \
                                                            (sheet['S' + str(j - 7)].value + sheet[
                                                                'S' + str(j - 8)].value +
                                                             sheet[
                                                                 'S' + str(j - 9)].value
                                                             + sheet['S' + str(j - 10)].value + sheet[
                                                                 'S' + str(j - 11)].value +
                                                             sheet[
                                                                 'S' + str(j - 12)].value)
                    if 'Изменение устьевого давления Pу при максимально допустимом режиме на ' + Entry18.get() + ' месяцев' in prognozi:
                        text_box.insert(END,
                                                    'Устьевое давление скважины ' + str(j - 57 - k) + ' при максимально допустимом режиме на конец ' + str(
                                                        k+1) + '-го месяца = ' + str(
                                                        sheet['S' + str(j)].value) + ' МПа\n')
                        #Давление пластовое при оптимальном режиме
            for k in range(0, mesyaci):
                for j in range(57 + k, 57 + i + k):
                    sheet['E' + str(j)].value = sheet['E' + str(j - 6)].value * (
                                            sheet['E' + str(j - 6)].value
                                            + sheet['E' + str(j - 5)].value +
                                            sheet[
                                                'E' + str(j - 4)].value +
                                            sheet[
                                                'E' + str(j - 3)].value
                                            + sheet['E' + str(j - 2)].value +
                                            sheet[
                                                'E' + str(j - 1)].value) / \
                                                            (sheet['E' + str(j - 7)].value + sheet[
                                                                'E' + str(j - 8)].value +
                                                             sheet[
                                                                 'E' + str(j - 9)].value
                                                             + sheet['E' + str(j - 10)].value + sheet[
                                                                 'E' + str(j - 11)].value +
                                                             sheet[
                                                                 'E' + str(j - 12)].value)
                    if 'Изменение пластового давления Pпл при оптимальном режиме на ' + Entry18.get() + ' месяцев' in prognozi:
                        text_box.insert(END,
                                                    'Пластовое давление скважины' + str(j - 57 - k+ 1) + ' при оптимальном режиме на конец ' + str(
                                                        k+1) + '-го месяца = ' + str(
                                                        sheet['E' + str(j)].value) + ' МПа\n')

        def hydrates_stvol(i):
            if 'Гидратообразование в стволе скважины' in ogranicheniya:
                text_box.insert(END, '\nГидраты в скважине образуются при температуре ' + str(round(T_ust[i-4], 2)) + '-' + str(round(T_bot[i-4], 2)) + ' C\n')
                text_box.insert(END, 'Текущая температура в стволе скважины ' + str(round(sheet['K' + str(i)].value, 2)) + '-' + str(39) + ' C\n')
        hydrates_stvol(i)

    predict_P()



disableds = [DISABLED,DISABLED,DISABLED,DISABLED,DISABLED,DISABLED,DISABLED,DISABLED,DISABLED,DISABLED,DISABLED,DISABLED,DISABLED,DISABLED,DISABLED,]
def change():
    n = int(Entry1.get())
    if n > 16:
        n = 16
    else:
        pass
    for i in range(0, n-1):
        Entries[i] = tk.Entry(group_3, bg="white", width=10, state=NORMAL)
        Entries[i].grid(row=(3+i), column=1, sticky=W)
        labelEntries[i] = tk.Label(group_3,font = 'Arial 9', text=("Между", str(i+1), "и", str(i+2), "скв.:"), state = NORMAL)
        labelEntries[i].grid(row=(3+i), column=0, sticky=W)
Entries =[]
labelEntries = []
def disable_water():
    checkbutton4 = tk.Checkbutton(group_2, text='Макс. безводный дебит', font='Arial 10', onvalue=1, offvalue=0,
                                  variable=intvar_cb4, state = NORMAL)
    checkbutton4.grid(row=0, column=0, sticky=W)
    checkbutton5 = tk.Checkbutton(group_2, text='Мин. дебит для выноса воды', font='Arial 10', onvalue=3, offvalue=2,
                                  variable=intvar_cb5, state=NORMAL)
    checkbutton5.grid(row=1, column=0, sticky=W)
    checkbutton25 = tk.Checkbutton(group_5, font='Arial 10', text='Дебит воды, Qв', onvalue=41, offvalue=40,
                                   variable=intvar_cb25, state=NORMAL)
    checkbutton25.grid(row=13, column=0, sticky=W)

####### ОКНО #######

window = tk.Tk()
#window.resizable(width=FALSE, height=FALSE)
window.title("ver 1.1")
window.geometry('945x720')

####### ВКЛАДКИ МЕНЮ #######

menu = Menu(window)
new_item = Menu(menu)
new_item.add_command(label='Новый')
new_item.add_separator()
new_item.add_command(label='Изменить')
menu.add_cascade(label='Файл', menu=new_item)

new_item2 = Menu(menu)
new_item2.add_command(label='Новый')
new_item2.add_separator()
new_item2.add_command(label='Изменить')
menu.add_cascade(label='Помощь', menu=new_item2)

CheckVar1 = IntVar()

####### ТЕКСТБОКС #######

group_6= tk.Label(padx=15, pady=10, text="")
group_6.place(x = 630, y = 675)
text_box = tk.Text(width=116, height=17, font = 'Arial 11')
text_box.place(x = 5, y = 400)

####### ТИП МЕСТОРОЖДЕНИЯ #######

var = IntVar()
def change1():
    checkbutton6 = tk.Checkbutton(group_2, text='Мин. дебит для выноса конденсата', font='Arial 10', onvalue=5,
                                  offvalue=4, variable=intvar_cb6, state=DISABLED)
    checkbutton6.grid(row=2, column=0, sticky=W)

    checkbutton7 = tk.Checkbutton(group_2, text='Огр. скорости по коррозии', font='Arial 10', onvalue=7, offvalue=6,
                                  variable=intvar_cb7, state=NORMAL)
    checkbutton7.grid(row=3, column=0, sticky=W)

    checkbutton8 = tk.Checkbutton(group_2, text='Гидратообр. на забое', font='Arial 10', onvalue=9, offvalue=8,
                                  variable=intvar_cb8, state=NORMAL)
    checkbutton8.grid(row=4, column=0, sticky=W)

    checkbutton9 = tk.Checkbutton(group_2, text='Гидратообр. на устье', font='Arial 10', onvalue=11, offvalue=10,
                                  variable=intvar_cb9, state=NORMAL)
    checkbutton9.grid(row=5, column=0, sticky=W)

    checkbutton10 = tk.Checkbutton(group_2, text='Гидратообр. в стволе скважины', font='Arial 10', onvalue=13,
                                   offvalue=12, variable=intvar_cb10, state=NORMAL)
    checkbutton10.grid(row=6, column=0, sticky=W)

    checkbutton11 = tk.Checkbutton(group_2, text='Гидратообр. в шлейфе', font='Arial 10', onvalue=15, offvalue=14,
                                   variable=intvar_cb11, state=NORMAL)
    checkbutton11.grid(row=7, column=0, sticky=W)

    checkbutton12 = tk.Checkbutton(group_2, text='По скорости в шлейфе', font='Arial 10', onvalue=17, offvalue=16,
                                   variable=intvar_cb12, state=NORMAL)
    checkbutton12.grid(row=8, column=0, sticky=W)
def change2():
    checkbutton6 = tk.Checkbutton(group_2, text='Мин. дебит для выноса конденсата', font='Arial 10', onvalue=5,
                                  offvalue=4, variable=intvar_cb6, state=NORMAL)
    checkbutton6.grid(row=2, column=0, sticky=W)

    checkbutton7 = tk.Checkbutton(group_2, text='Огр. скорости по коррозии', font='Arial 10', onvalue=7, offvalue=6,
                                  variable=intvar_cb7, state=NORMAL)
    checkbutton7.grid(row=3, column=0, sticky=W)

    checkbutton8 = tk.Checkbutton(group_2, text='Гидратообр. на забое', font='Arial 10', onvalue=9, offvalue=8,
                                  variable=intvar_cb8, state=NORMAL)
    checkbutton8.grid(row=4, column=0, sticky=W)

    checkbutton9 = tk.Checkbutton(group_2, text='Гидратообр. на устье', font='Arial 10', onvalue=11, offvalue=10,
                                  variable=intvar_cb9, state=NORMAL)
    checkbutton9.grid(row=5, column=0, sticky=W)

    checkbutton10 = tk.Checkbutton(group_2, text='Гидратообр. в стволе скважины', font='Arial 10', onvalue=13,
                                   offvalue=12, variable=intvar_cb10, state=NORMAL)
    checkbutton10.grid(row=6, column=0, sticky=W)

    checkbutton11 = tk.Checkbutton(group_2, text='Гидратообр. в шлейфе', font='Arial 10', onvalue=15, offvalue=14,
                                   variable=intvar_cb11, state=NORMAL)
    checkbutton11.grid(row=7, column=0, sticky=W)

    checkbutton12 = tk.Checkbutton(group_2, text='По скорости в шлейфе', font='Arial 10', onvalue=17, offvalue=16,
                                   variable=intvar_cb12, state=NORMAL)
    checkbutton12.grid(row=8, column=0, sticky=W)
def change3():
    checkbutton6 = tk.Checkbutton(group_2, text='Мин. дебит для выноса конденсата', font='Arial 10', onvalue=5,
                                  offvalue=4, variable=intvar_cb6, state=DISABLED)
    checkbutton6.grid(row=2, column=0, sticky=W)

    checkbutton7 = tk.Checkbutton(group_2, text='Огр. скорости по коррозии', font='Arial 10', onvalue=7, offvalue=6,
                                  variable=intvar_cb7, state=DISABLED)
    checkbutton7.grid(row=3, column=0, sticky=W)

    checkbutton8 = tk.Checkbutton(group_2, text='Гидратообр. на забое', font='Arial 10', onvalue=9, offvalue=8,
                                  variable=intvar_cb8, state=DISABLED)
    checkbutton8.grid(row=4, column=0, sticky=W)

    checkbutton9 = tk.Checkbutton(group_2, text='Гидратообр. на устье', font='Arial 10', onvalue=11, offvalue=10,
                                  variable=intvar_cb9, state=DISABLED)
    checkbutton9.grid(row=5, column=0, sticky=W)

    checkbutton10 = tk.Checkbutton(group_2, text='Гидратообр. в стволе скважины', font='Arial 10', onvalue=13,
                                   offvalue=12, variable=intvar_cb10, state=DISABLED)
    checkbutton10.grid(row=6, column=0, sticky=W)

    checkbutton11 = tk.Checkbutton(group_2, text='Гидратообр. в шлейфе', font='Arial 10', onvalue=15, offvalue=14,
                                   variable=intvar_cb11, state=DISABLED)
    checkbutton11.grid(row=7, column=0, sticky=W)

    checkbutton12 = tk.Checkbutton(group_2, text='По скорости в шлейфе', font='Arial 10', onvalue=17, offvalue=16,
                                   variable=intvar_cb12, state=NORMAL)
    checkbutton12.grid(row=8, column=0, sticky=W)

disableds_type = [DISABLED,DISABLED,DISABLED,DISABLED,DISABLED,DISABLED,DISABLED]

group_1 = tk.LabelFrame(padx=15, pady=5, font = 'Arial 9', text="Тип месторождения:")
Radiobut1 = tk.Radiobutton(group_1, font = 'Arial 10', text='Газовое', value=0, variable=var,command = change1)
Radiobut1.grid(row = 0, column = 0, sticky = W)

Radiobut2 = tk.Radiobutton(group_1, font = 'Arial 10', text='Газоконденсатное', value=1,variable=var,command = change2)
Radiobut2.grid(row = 1, column = 0, sticky = W)

Radiobut3 = tk.Radiobutton(group_1, font = 'Arial 10', text='Нефтяное', value=2, variable=var,command = change3)
Radiobut3.grid(row = 2, column = 0, sticky = W)



intvar_cb38 = IntVar()
checkbutton_voda = tk.Checkbutton(group_1, text = 'Наличие подошвенной воды     ', font = 'Arial 10', onvalue = 63, offvalue = 62, variable = intvar_cb38, command = disable_water)
checkbutton_voda.grid(row = 3, column = 0,sticky = W)
group_1.place(x=5, y=1)


####### ОГРАНИЧЕНИЯ РЕЖИМА #######
#intvars:
intvar_cb4 = IntVar()
intvar_cb5 = IntVar()
intvar_cb6 = IntVar()
intvar_cb7 = IntVar()
intvar_cb8 = IntVar()
intvar_cb9 = IntVar()
intvar_cb10 = IntVar()
intvar_cb11 = IntVar()
intvar_cb12 = IntVar()

group_2= tk.LabelFrame(padx=5, pady=5, font = 'Arial 10', text="Ограничения:")
checkbutton4 = tk.Checkbutton(group_2, text = 'Макс. безводный дебит', font = 'Arial 10', onvalue = 1, offvalue = 0, variable = intvar_cb4, state = DISABLED)
checkbutton4.grid(row = 0, column = 0,sticky = W)

checkbutton5 = tk.Checkbutton(group_2, text = 'Мин. дебит для выноса воды', font = 'Arial 10', onvalue = 3, offvalue = 2, variable = intvar_cb5, state = DISABLED)
checkbutton5.grid(row = 1, column = 0,sticky = W)

checkbutton6 = tk.Checkbutton(group_2, text = 'Мин. дебит для выноса конденсата', font = 'Arial 10', onvalue = 5, offvalue = 4, variable = intvar_cb6, state = disableds_type[0])
checkbutton6.grid(row = 2, column = 0,sticky = W)

checkbutton7 = tk.Checkbutton(group_2, text = 'Огр. скорости по коррозии', font = 'Arial 10', onvalue = 7, offvalue = 6, variable = intvar_cb7, state = disableds_type[1])
checkbutton7.grid(row = 3, column = 0,sticky = W)

checkbutton8 = tk.Checkbutton(group_2, text = 'Гидратообр. на забое', font = 'Arial 10', onvalue = 9, offvalue = 8, variable = intvar_cb8, state = disableds_type[2])
checkbutton8.grid(row = 4, column = 0,sticky = W)

checkbutton9 = tk.Checkbutton(group_2, text = 'Гидратообр. на устье', font = 'Arial 10', onvalue = 11, offvalue = 10, variable = intvar_cb9, state = disableds_type[3])
checkbutton9.grid(row = 5, column = 0,sticky = W)

checkbutton10 = tk.Checkbutton(group_2, text = 'Гидратообр. в стволе скважины', font = 'Arial 10', onvalue = 13, offvalue = 12, variable = intvar_cb10, state = disableds_type[4])
checkbutton10.grid(row = 6, column = 0,sticky = W)

checkbutton11 = tk.Checkbutton(group_2, text = 'Гидратообр. в шлейфе', font = 'Arial 10', onvalue = 15, offvalue = 14, variable = intvar_cb11, state = disableds_type[5])
checkbutton11.grid(row = 7, column = 0,sticky = W)

checkbutton12 = tk.Checkbutton(group_2, text = 'По скорости в шлейфе', font = 'Arial 10', onvalue = 17, offvalue = 16, variable = intvar_cb12, state = disableds_type[6])
checkbutton12.grid(row = 8, column = 0,sticky = W)
group_2.place(x = 5, y = 135)


####### СКВАЖИНЫ #######
group_3 = tk.LabelFrame(padx=15, pady=4, text="")
label1 = tk.Label(group_3,font = 'Arial 9', text="Кол-во скважин:")
label1.grid(row = 0, column = 0,)

Entry1 = tk.Entry(group_3, width=15, textvariable = tk.IntVar)
Entries.append(Entry1)
Entry1.grid(row=1, column=0, )

But1 = tk.Button(group_3, text = 'Применить', font = 'Arial 7', command = change)
But1.grid(row = 1, column = 1,sticky = W, padx = 5)

label2 = tk.Label(group_3,font = 'Arial 9', text="Расстояния, м:")
label2.grid(row = 2, column = 0,)

Entry2 = tk.Entry(group_3, bg="white", width=10, state = disableds[0])
Entries.append(Entry2)
Entry2.grid(row = 3, column = 1, sticky = W)
labelEntry2 = tk.Label(group_3,font = 'Arial 9', text="Между 1 и 2 скв.:", state = disableds[0])
labelEntries.append(labelEntry2)
labelEntry2.grid(row = 3, column = 0,sticky = W)

Entry3 = tk.Entry(group_3, bg="white", width=10, state = disableds[1])
Entries.append(Entry3)
Entry3.grid(row = 4, column = 1, sticky = W)
labelEntry3 = tk.Label(group_3,font = 'Arial 9', text="Между 2 и 3 скв.:", state = disableds[1])
labelEntries.append(labelEntry3)
labelEntry3.grid(row = 4, column = 0,sticky = W)

Entry4 = tk.Entry(group_3, bg="white", width=10, state = disableds[2])
Entries.append(Entry4)
Entry4.grid(row = 5, column = 1, sticky = W)
labelEntry4 = tk.Label(group_3,font = 'Arial 9', text="Между 3 и 4 скв.:", state = disableds[2])
labelEntries.append(labelEntry4)
labelEntry4.grid(row = 5, column = 0,sticky = W)

Entry5 = tk.Entry(group_3, bg="white", width=10, state = disableds[3])
Entries.append(Entry5)
Entry5.grid(row = 6, column = 1, sticky = W)
labelEntry5 = tk.Label(group_3,font = 'Arial 9', text="Между 4 и 5 скв.:", state = disableds[3])
labelEntries.append(labelEntry5)
labelEntry5.grid(row = 6, column = 0,sticky = W)

Entry6 = tk.Entry(group_3, bg="white", width=10, state = disableds[4])
Entries.append(Entry6)
Entry6.grid(row = 7, column = 1, sticky = W)
labelEntry6 = tk.Label(group_3,font = 'Arial 9', text="Между 5 и 6 скв.:", state = disableds[4])
labelEntries.append(labelEntry6)
labelEntry6.grid(row = 7, column = 0,sticky = W)

Entry7 = tk.Entry(group_3, bg="white", width=10, state = disableds[5])
Entries.append(Entry7)
Entry7.grid(row = 8, column = 1, sticky = W)
labelEntry7 = tk.Label(group_3,font = 'Arial 9', text="Между 6 и 7 скв.:", state = disableds[5])
labelEntries.append(labelEntry7)
labelEntry7.grid(row = 8, column = 0,sticky = W)

Entry8 = tk.Entry(group_3, bg="white", width=10, state = disableds[6])
Entries.append(Entry8)
Entry8.grid(row = 9, column = 1, sticky = W)
labelEntry8 = tk.Label(group_3,font = 'Arial 9', text="Между 7 и 8 скв.:", state = disableds[6])
labelEntries.append(labelEntry8)
labelEntry8.grid(row = 9, column = 0,sticky = W)

Entry9 = tk.Entry(group_3, bg="white", width=10, state = disableds[7])
Entries.append(Entry9)
Entry9.grid(row = 10, column = 1, sticky = W)
labelEntry9 = tk.Label(group_3,font = 'Arial 9', text="Между 8 и 9 скв.:", state = disableds[7])
labelEntries.append(labelEntry9)
labelEntry9.grid(row = 10, column = 0,sticky = W)

Entry10 = tk.Entry(group_3, bg="white", width=10, state = disableds[8])
Entries.append(Entry10)
Entry10.grid(row = 11, column = 1, sticky = W)
labelEntry10 = tk.Label(group_3,font = 'Arial 9', text="Между 9 и 10 скв.:", state = disableds[8])
labelEntries.append(labelEntry10)
labelEntry10.grid(row = 11, column = 0,sticky = W)

Entry11 = tk.Entry(group_3, bg="white", width=10, state = disableds[9])
Entries.append(Entry11)
Entry11.grid(row = 12, column = 1, sticky = W)
labelEntry11 = tk.Label(group_3,font = 'Arial 9', text="Между 10 и 11 скв.:", state = disableds[9])
labelEntries.append(labelEntry11)
labelEntry11.grid(row = 12, column = 0,sticky = W)

Entry12 = tk.Entry(group_3, bg="white", width=10, state = disableds[10])
Entries.append(Entry12)
Entry12.grid(row = 13, column = 1, sticky = W)
labelEntry12 = tk.Label(group_3,font = 'Arial 9', text="Между 11 и 12 скв.:", state = disableds[10])
labelEntries.append(labelEntry12)
labelEntry12.grid(row = 13, column = 0,sticky = W)

Entry13 = tk.Entry(group_3, bg="white", width=10, state = disableds[11])
Entries.append(Entry13)
Entry13.grid(row = 14, column = 1, sticky = W)
labelEntry13 = tk.Label(group_3,font = 'Arial 9', text="Между 12 и 13 скв.:", state = disableds[11])
labelEntries.append(labelEntry13)
labelEntry13.grid(row = 14, column = 0,sticky = W)

Entry14 = tk.Entry(group_3, bg="white", width=10, state = disableds[12])
Entries.append(Entry14)
Entry14.grid(row = 15, column = 1, sticky = W)
labelEntry14 = tk.Label(group_3,font = 'Arial 9', text="Между 13 и 14 скв.:", state = disableds[12])
labelEntries.append(labelEntry14)
labelEntry14.grid(row = 15, column = 0,sticky = W)

Entry15 = tk.Entry(group_3, bg="white", width=10, state = disableds[13])
Entries.append(Entry15)
Entry15.grid(row = 16, column = 1, sticky = W)
labelEntry15 = tk.Label(group_3,font = 'Arial 9', text="Между 14 и 15 скв.:", state = disableds[13])
labelEntries.append(labelEntry15)
labelEntry15.grid(row = 16, column = 0,sticky = W)

Entry16 = tk.Entry(group_3, bg="white", width=10, state = disableds[14])
Entries.append(Entry16)
Entry16.grid(row = 17, column = 1, sticky = W)
labelEntry16 = tk.Label(group_3,font = 'Arial 9', text="Между 15 и 16 скв.:", state = disableds[14])
labelEntries.append(labelEntry16)
labelEntry16.grid(row = 17, column = 0,sticky = W)
group_3.place(x = 265, y = 9)


####### УЧЕТ ПАРАМЕТРОВ #######

#intvars:
intvar_cb14 = IntVar()
intvar_cb15 = IntVar()
intvar_cb16 = IntVar()
intvar_cb17 = IntVar()
intvar_cb18 = IntVar()
intvar_cb19 = IntVar()
intvar_cb20 = IntVar()
intvar_cb21 = IntVar()
intvar_cb22 = IntVar()
intvar_cb23 = IntVar()
intvar_cb24 = IntVar()
intvar_cb25 = IntVar()

group_5 = tk.LabelFrame(padx=1, pady=25, text="")
tk.Label(group_5,font = 'Arial 10', text = 'Учитывать параметры:').grid(row = 0, column = 0, sticky = S)
tk.Label(group_5,font = 'Arial 10', text = '').grid(row = 2, column = 0)

checkbutton14 = tk.Checkbutton(group_5,font = 'Arial 10', text = 'Статическое давление, Рст', onvalue = 19, offvalue = 18, variable = intvar_cb14)
checkbutton14.grid(row = 2, column = 0,sticky = W)

checkbutton15 = tk.Checkbutton(group_5,font = 'Arial 10', text = 'Пластовое давление, Рпл', onvalue = 21, offvalue = 20, variable = intvar_cb15)
checkbutton15.grid(row = 3, column = 0,sticky = W)

checkbutton16 = tk.Checkbutton(group_5,font = 'Arial 10', text = 'Рабочий дебит, Q', onvalue = 23, offvalue = 22, variable = intvar_cb16)
checkbutton16.grid(row = 4, column = 0,sticky = W)

checkbutton17 = tk.Checkbutton(group_5,font = 'Arial 10', text = 'Депрессия, ΔP', onvalue = 25, offvalue = 24, variable = intvar_cb17)
checkbutton17.grid(row = 5, column = 0,sticky = W)

checkbutton18 = tk.Checkbutton(group_5,font = 'Arial 10', text = 'Устьевое давление, Pус', onvalue = 27, offvalue = 26, variable = intvar_cb18)
checkbutton18.grid(row = 6, column = 0,sticky = W)

checkbutton19 = tk.Checkbutton(group_5,font = 'Arial 10', text = 'Затрубное давление, Рзт', onvalue = 29, offvalue = 28, variable = intvar_cb19)
checkbutton19.grid(row = 7, column = 0,sticky = W)

checkbutton20 = tk.Checkbutton(group_5,font = 'Arial 10', text = 'Давление в шлейфе, Ршл', onvalue = 31, offvalue = 30, variable = intvar_cb20)
checkbutton20.grid(row = 8, column = 0,sticky = W)

checkbutton21 = tk.Checkbutton(group_5,font = 'Arial 10', text = 'Устьевая температура, Тус', onvalue = 33, offvalue = 32, variable = intvar_cb21)
checkbutton21.grid(row = 9, column = 0,sticky = W)

checkbutton22 = tk.Checkbutton(group_5,font = 'Arial 10', text = 'Межколонное давление, Рмк', onvalue = 35, offvalue = 34, variable = intvar_cb22)
checkbutton22.grid(row = 10, column = 0,sticky = W)

checkbutton23 = tk.Checkbutton(group_5,font = 'Arial 10', text = 'Давление на входе УКПГ, Рвх', onvalue = 37, offvalue = 36, variable = intvar_cb23)
checkbutton23.grid(row = 11, column = 0,sticky = W)

checkbutton24 = tk.Checkbutton(group_5,font = 'Arial 10', text = 'Температура на входе УКПГ, Твх', onvalue = 39, offvalue = 38, variable = intvar_cb24)
checkbutton24.grid(row = 12, column = 0,sticky = W)

checkbutton25 = tk.Checkbutton(group_5,font = 'Arial 10', text = 'Дебит воды, Qв', onvalue = 41, offvalue = 40, variable = intvar_cb25, state = DISABLED)
checkbutton25.grid(row = 13, column = 0,sticky = W)
group_5.place(x = 485, y = 10)

####### ИНИЦИАЛИЗАЦИЯ ВЕСОВ #######
group_7= tk.LabelFrame(padx=5, pady=4,font = 'Arial 10', text="Инициализация весов")
label4 = tk.Label(group_7,font = 'Arial 10', text="Количество итераций (min. 100):").grid(row = 0, column = 0,sticky = W)
Entry17 = tk.Entry(group_7, bg="white", width=21)
Entry17.grid(row = 1, column = 0,sticky = S)
group_7.place(x = 725, y = 2)

####### ПРОГНОЗНЫЕ ПОКАЗАТЕЛИ #######

#intvars
intvar_cb27 = IntVar()
intvar_cb28 = IntVar()
intvar_cb29 = IntVar()
intvar_cb30 = IntVar()
intvar_cb31 = IntVar()
intvar_cb32 = IntVar()
intvar_cb33 = IntVar()
intvar_cb34 = IntVar()
intvar_cb35 = IntVar()

group_8= tk.LabelFrame(padx=5, pady=3,font = 'Arial 10', text="Прогнозные показатели")
#label5 = tk.Label(group_8, text="Сделать прогноз по:").grid(row = 0, column = 0)
checkbutton27 = tk.Checkbutton(group_8,font = 'Arial 10', text = 'Qнак. по макс. доп. режиму', onvalue = 43, offvalue = 42, variable = intvar_cb27)
checkbutton27.grid(row = 1, column = 0,sticky = W)

checkbutton28 = tk.Checkbutton(group_8,font = 'Arial 10', text = 'Qнак. по мин. доп. режиму', onvalue = 45, offvalue = 44, variable = intvar_cb28)
checkbutton28.grid(row = 2, column = 0,sticky = W)

checkbutton29 = tk.Checkbutton(group_8,font = 'Arial 10', text = 'Qнак. по оптим. режиму', onvalue = 47, offvalue = 46, variable = intvar_cb29)
checkbutton29.grid(row = 3, column = 0,sticky = W)

checkbutton30 = tk.Checkbutton(group_8,font = 'Arial 10', text = 'Pпл при макс. доп. режиме', onvalue = 49, offvalue = 48, variable = intvar_cb30)
checkbutton30.grid(row = 4, column = 0,sticky = W)

checkbutton31 = tk.Checkbutton(group_8,font = 'Arial 10', text = 'Pпл при мин. доп. режиме', onvalue = 51, offvalue = 50, variable = intvar_cb31)
checkbutton31.grid(row = 5, column = 0,sticky = W)

checkbutton32 = tk.Checkbutton(group_8,font = 'Arial 10', text = 'Рпл при оптим. режиме', onvalue = 53, offvalue = 52, variable = intvar_cb32)
checkbutton32.grid(row = 6, column = 0,sticky = W)

checkbutton33 = tk.Checkbutton(group_8,font = 'Arial 10', text = 'Pу при макс. доп. режиме', onvalue = 55, offvalue = 54, variable = intvar_cb33)
checkbutton33.grid(row = 7, column = 0,sticky = W)

checkbutton34 = tk.Checkbutton(group_8,font = 'Arial 10', text = 'Pу при мин. доп. режиме', onvalue = 57, offvalue = 56, variable = intvar_cb34)
checkbutton34.grid(row = 8, column = 0,sticky = W)

checkbutton35 = tk.Checkbutton(group_8,font = 'Arial 10', text = 'Ру при оптим. режиме', onvalue = 59, offvalue = 58, variable = intvar_cb35)
checkbutton35.grid(row = 9, column = 0,sticky = W)

label6 = tk.Label(group_8,font = 'Arial 10', text="Количество месяцев прогноза:").grid(row = 10, column = 0,sticky = W)
Entry18 = tk.Entry(group_8, bg="white", width=21)
Entry18.grid(row = 11, column = 0,sticky = S)
group_8.place(x = 725, y = 72)

####### КНОПКА #######
But2 = tk.Button(text = 'Загрузить данные для расчета',font = 'Arial 10', command = rasschet_final).place(x = 730, y = 372)

####### ЭКСПОРТ #######

intvar_cb36 = IntVar()
intvar_cb37 = IntVar()
tk.Label(group_6, text = 'Экспортировать результат как:').grid(row = 0, column = 0,sticky = E)
checkbutton36 = tk.Checkbutton(group_6, onvalue = 59, offvalue = 58, variable = intvar_cb36, text = '.txt').grid(row = 0, column = 1,sticky = E)
checkbutton37 = tk.Checkbutton(group_6, onvalue = 61, offvalue = 60, variable = intvar_cb37, text = '.xlsx').grid(row = 0, column = 2,sticky = E)
tk.Button(group_6, text = 'ОК').grid(row = 0, column = 3,sticky = E)
group_6.place(x = 625, y = 689)

window.config(menu=menu)
window.mainloop()

