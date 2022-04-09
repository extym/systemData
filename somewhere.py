# from openpyxl import Workbook
# from os.path import exists
# from server_script import dict_data
# import pandas as pd
# # #
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
# #
# # df = pd.DataFrame(dict_data, index=[0, 1, 2, 3, 4, 5, 6 ,7])
# # while dict_data is not 0:
# #     df.to_excel('./test.xlsx')
# row = 0
# sheet = 1
# # df = pd.DataFrame.from_dict(data=dict_data, orient='columns')
# #df = df.T
# for dict in dict_data:
#     dict = pd.DataFrame.from_dict(data=dict_data, orient='index',
#                                   columns=['pid', 'ppid', 'name', 'create_time', 'ID', 'ip_address', 'port'])
#     row += 1
#     print (dict)
#     dict.to_excel('dict1.xlsx',sheet_name=sheet, startrow=row)
