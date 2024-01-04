import socket

def send_card_number_to_mac(card_number, mac_ip, port):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.connect((mac_ip, port))
        s.sendall(card_number.encode())

def read_card():
    # Placeholder for card reading logic
    return input("Enter card number: ")

while True:
    card_number = read_card()
    send_card_number_to_mac(card_number, 'MAC_IP_ADDRESS', 12345)  # Replace with actual Mac IP and port
