# from openpyxl import Workbook
# from os.path import exists
# from server_script import dict_data
# import csv
# #
# # # Теоретически правильный алгоритм:
# # # 1. Клиент отправляет запрос и получает данные о процессе и всех процессах
# # # 2. Пишется в таблицу Имя и хеш процесса на основе совокупности его признаков
# # # 3. Полученная информация о процессах пишется в таблицу Мониторинг
# # # 4. Данные из таблицы Мониторинг интерпретируются согласно скрипту и пишутся в таблицу Статистика
# #
#
# #file_exists = exists("monitoring.xlsx")
# import xlsxwriter
#
# d = dict_data
# print(type(d))
#
# with open('../monitor.csv', newline='') as csvfile:
#     spamreader = csv.reader(csvfile, delimiter=';')
#     for row in spamreader:
#         newobj = dict(spamreader[0], spamreader[1:6])
#         print(type(newobj), newobj)