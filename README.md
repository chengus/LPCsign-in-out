# LPC sign in/out system
LPC sign-in/out system at guard house with RFID, Tkinter, Raspberry Pi, Python, Excel, and Google Drive API.

## Hardware setup
1. Mac/Windows computer: internet access, two monitor ports, ethernet(preferred)/wifi connectivity
2. Raspberry Pi (ethernet preferred) - Assigned static IP/Direct Connection
3. RFID card reader
4. Two monitors

Main computer is connected to RPi through ethernet/wifi.
RFID reader(s) connected directly to RPi.
Both monitors are connected to the main computer.

## Software Logic
Starting script initialize main script & parameter -> main script starts listening indefinately
RPi capture RFID reader input indefinately -> sent to main script with socket -> process info -> display info -> record data 

## Future plans
- Package tkinter app to .exe
- Color code in `tksheets` & exported Excel/Sheets
- Incorporate web-based dashboard to assign overnights/extensions/gating
- Upload to specific folder on google drive
- Improve UI 


