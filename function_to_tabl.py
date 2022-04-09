#
# from server_script import dict_data # ---- входные данные для моего скрипта (в )
# import openpyxl
# import pandas as pd
#
# read_file = pd.read_csv('realytrac.csv')
# read_file.to_excel('monitoring.xlsx', index=None, header=True)
#
# def calculate_max_rows(sheets, column):
#     ''' Вычисление количества непустых строк на странице
#     ВНИМАНИЕ на столбец, по которому происходит просмотр макимального количества строк'''
#     not_empty_rows = sheets.max_row  # непустые строки
#     # вычисляю номер последней строки, содержащей URL
#     while not_empty_rows > 1:
#         if sheets.cell(row=not_empty_rows, column=column).value == None:  # ----- проверка по строке D -----
#             not_empty_rows -= 1
#         else:
#             break
#
#     return not_empty_rows
#
# def time_means():
#     '''вычислим среднее время  процесса'''
#     sum_ = 0
#     time_m = 0
#     row_summ = 0
#     last_row_summ = calculate_max_rows(ws1, 2)
#     for item_row in range(last_row_summ):
#         row_summ = item_row + 2
#         if ws1.cell(row=row_summ, column=2).value != None:
#             sum_ += ws1.cell(row=row_summ, column=2).value
#         else: break
#         ws1.cell(row=row_summ, column=8).value = round(sum_ / (row_summ - 1),2)
#
#     return time_m
#
# def cpc_means():
#     '''вычислим среднее количество копий процесса'''
#     sum_ = 0
#     cpc_m = 0
#     row_summ = 0
#     last_row_summ = calculate_max_rows(ws1, 2)
#     for item_row in range(last_row_summ):
#         row_summ = item_row + 2
#
#         if ws1.cell(row=row_summ, column=6).value != None:
#             sum_ += ws1.cell(row=row_summ, column=6).value
#         else: break
#         ws1.cell(row=row_summ, column=7).value = round(sum_ / (row_summ - 1),2)
#
#     return cpc_m
#
# #
# #     # ----- генерируем словарь процесса --------
# #     k_dict = {'create_time': time_r,
# #               'IP': ip_r,
# #               'PORT': port_r,
# #               'ppid': ppid_r,
# #               'count_proc_copy': cpc_r}
# #     return k_dict
#
# # ----- запись в табл  ------
#
#
# file_m = 'monitoring.xlsx'
#
# # import pandas as pd
# # # import
# # di = dict_data
# # z = pd.DataFrame([di])#, index=[0])
# # z.to_excel("file_name.xlsx")
# # # This is a sample Python script.
#
# wb_m = openpyxl.load_workbook(file_m, read_only=False)
# # ----- проверка на то, что эксель файл закрыт ---
# try:
#     wb_m.save(file_m)
# except PermissionError as ex:
#     print(f" Ошибка. файл {file_m} открыт. Закройте файл")
#     print(input('\n Для продолжения - нажмите "Enter"'))
#     exit()
#
#     # {'create_time': 163.77, 'IP': '289.35.62.215', 'PORT': 999, 'ppid': 2, 'count_proc_copy': 123}
#
# ws1 = wb_m['Лист1']
# last_row = calculate_max_rows(ws1, 2)
#
# # ---- организация цикла для внесения данных в таблицу Мониторинг------
# item_k = dict_data
#
#
# row_ = last_row + 1
# try:  ws1.cell(row=row_, column=1).value = int(ws1.cell(row=row_ - 1, column=1).value) + 1
# except: ws1.cell(row=row_, column=1).value = 1
# ws1.cell(row=row_, column=2).value = item_k['create_time']
# ws1.cell(row=row_, column=3).value = item_k['ip_address']
# ws1.cell(row=row_, column=4).value = item_k['port']
# ws1.cell(row=row_, column=5).value = item_k['ppid']
# ws1.cell(row=row_, column=6).value = item_k['count_proc_copy']
# ws1.cell(row=row_, column=7).value = item_k['ID']
# ws1.cell(row=row_, column=8).value = item_k['pid']
# ws1.cell(row=row_, column=9).value = item_k['name']
#
# wb_m.save(file_m)
# # ---- запуск функции расчета среднего времени процесса -----
# time_means()
# # ---- запуск функции расчета среднего количества копий процесса -----
# cpc_means()
#
# wb_m.save(file_m)
# wb_m.close()
