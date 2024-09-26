import socket

def send_card_number_to_pc(card_number, mac_ip, port):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.connect((mac_ip, port))
        s.sendall(card_number.encode())

def read_card():
    card_number = input("Enter card number: ")
    if not card_number.isdigit():
        raise ValueError("Card number must contain only digits (0-9).")
    return card_number

PCIP = input("Enter PC IP address: ")

while True:
    try:
        card_number = read_card()
        send_card_number_to_pc(card_number, PCIP, 12345)  # Replace with PC IP address
    except Exception as e:
        print(e)
        

