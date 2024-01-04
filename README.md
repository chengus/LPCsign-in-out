# LPC sign in/out system
LPC sign-in/out system at guard house with RFID, Tkinter, Raspberry Pi, Python, Excel, and Google Drive API.

## Hardware setup
1. Mac/Windows computer: internet access, two monitor ports, ethernet or wifi connectivity
2. Raspberry Pi (best with ethernet port) - Assigned static IP
3. RFID card reader (min 1)
4. Two monitors

Main computer is connected to RPi through ethernet/wifi.
RFID reader(s) connected directly to RPi.
Both monitors are connected to the main computer.

## Software setup - RPi
1. Run PiScript.py from terminal
2. Record IP address of main computer
3. Ensure python3 installed
In PiScript.py:

    `while True:
   
        card_number = read_card()
        send_card_number_to_mac(card_number, 'MAC_IP_ADDRESS', 12345)  # Replace with actual Mac IP and port`

Run `python3 /path/to/PiScript.py/` in terminal
