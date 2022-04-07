import psutil
import datetime
import socket
import uuid
import pandas as pd
import csv
import json

ip_Adrr = 0
ip_Port = 0

def getid():
    id = uuid.uuid4().hex  # setID
    return id

def char_replace(string):
    char_replace = {'(': '', ')': ''}
    for key, value in char_replace.items():
        string = string.replace(key, value)
        return string

def write_csv(data):
  with open('realytrac.csv', 'a') as f:
    writer = csv.writer(f, delimiter=';')
    writer.writerow((data['pid'],
                    data['name'],
                    data['ip_address'],
                    data['port'],
                    data['ppid'],
                    data['ID'],
                    data['create_time']))

def get_ip(connection):
    if connection is not None and len(connection) > 0:
        connect = str(connection[0])
        connect_ = connect.replace('pconn', '')
        connectlast = connect_.split(', ')
        for conn_last in connectlast:
            if conn_last.startswith('raddr=addr'):
                temp_conn = conn_last.split('=')
                ip_Adrr = temp_conn[2]
                return ip_Adrr

def get_port(connection):
    if connection is not None and len(connection) > 0:
        connect = str(connection[0])
        connect_ = connect.replace('pconn', '')
        connectlast = connect_.split(', ')
        for conn_last in connectlast:
            if conn_last.startswith('raddr=addr') is not '' and conn_last.startswith('port'):
                temp_conn = conn_last.split('=')
                ip_Port = temp_conn[1].replace(')', '')
                return ip_Port

# def write_xlsx(data):
#     i = 0
#     for data in data:
#         i += i
#     writy = pd.DataFrame(data, index=i)
#     writy.to_excel("file_name.xlsx")

# Время
def timeFor(timestamp):
  id_time = datetime.datetime.fromtimestamp(timestamp)
  return id_time

# Получить текущее системное время
current_time = datetime.datetime.now().strftime("%F %T") #% F год месяц день% T час минута секунда
# Get IP address
s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
s.connect(("8.8.8.8", 80))
ip_adrr = str(s.getsockname()[0]).split('.')
vm_ip = s.getsockname()[0]
host_ip = ip_adrr[0]+'.'+ ip_adrr[1]+'.' + ip_adrr[2] + '.1'
print("Host IP: ", host_ip)
print("VM IP", vm_ip)
s.close()

sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)  # создаем сокет
sock.bind(('', 55000))  # связываем сокет с портом, где он будет ожидать сообщения
sock.listen(10)  # указываем сколько может сокет принимать соединений
print('Server is running, please, press ctrl+c to stop')
while True:
    conn, addr = sock.accept()  # начинаем принимать соединения
    print('connected:', addr)  # выводим информацию о подключении
    data = conn.recv(1024)  # принимаем данные от клиента, по 1024 байт
    print(str(data))

    for proc in psutil.process_iter(['pid', 'name']):
        sp = proc.info
        # id = uuid.uuid4().hex  # setID
        pid = sp['pid']
        p = psutil.Process(pid)
        global dict_data
        dict_data = p.as_dict(attrs=['pid', 'name', 'connections', 'create_time', 'ppid'])
        dict_data['ID'] = getid()
        create_time = dict_data['create_time']
        time = timeFor(dict_data['create_time'])
        dict_data['create_time'] = time
        connect = dict_data['connections'] #??
        dict_data['ip_address'] = get_ip(connect)
        dict_data['port'] = get_port(connect)
        dict_data.pop('connections')
        write_csv(dict_data)
        print(dict_data)
    conn.send(bytes('Ok', encoding='utf-8'))  # в ответ клиенту отправляем сообщение в верхнем регистре
#conn.close()  # закрываем соединение
#
# read_file = pd.read_csv('realytrac.csv')
# read_file.to_excel('monitoring.xlsx', index=None, header=True)
# Press the green button in the gutter to run the script.
# if __name__ == '__main__':
#     print_hi('PyCharm')


