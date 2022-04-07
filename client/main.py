import socket
import datetime
def monitor(time):
    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)  # создаем сокет
    sock.connect(('localhost', 55000))  # подключемся к серверному сокету
    sock.send(bytes('Hello', encoding = 'UTF-8'))  # отправляем сообщение
    data = sock.recv(1024)  # читаем ответ от серверного сокета
    print(data)

def main():
  while True:
    monitor(5)  # Выполнять каждые 5 секунд


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    main()