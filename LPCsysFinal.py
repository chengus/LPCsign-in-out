import pandas as pd
import tkinter as tk
from tkinter import scrolledtext, ttk, messagebox, Toplevel, Button, filedialog
from pandas import ExcelWriter
import sqlite3, datetime, os, tempfile, sys, subprocess, shutil, tksheet, socket, threading
from PIL import Image, ImageTk
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import os.path
import pickle
from googleapiclient.http import MediaFileUpload

#TODO: Add export to excel button V
#TODO: Add ask to save before close V
#TODO: Add curfew V
#TODO: Add colorcoding to late returns V
#TODO: Add multimonitor support V
#TODO: Add starting screen: port, student xlsx path, montitor selection V
#TODO: Add school logo V
#TODO: Google drive integration V
#TODO: Remove scrollbar V

guard_mon = sys.argv[1]
student_mon = sys.argv[2]
port = sys.argv[3]
student_xlsx_path = sys.argv[4]


allowedLeaveTime = 600
allowedReturnTimeWeekday = 2100
allowedReturnTimeWeekend = 2300
bufferTime = 10
fontSize = 40

SCOPES = ['https://www.googleapis.com/auth/drive']

def parse_monitor_info(info):
    # Exmp. format: "Monitor 0 (1512x982)"
    parts = info.split(' ')
    monitor_id = int(parts[1]) 
    resolution_part = parts[2].strip('()')
    width, height = map(int, resolution_part.split('x'))
    return monitor_id, width, height

guard_mon_id, guard_monitor_width, guard_monitor_height = parse_monitor_info(guard_mon)
student_mon_id, student_monitor_width, student_monitor_height = parse_monitor_info(student_mon)

"""
print(guard_monitor_height)
print(guard_monitor_width)

print(student_monitor_height)
print(student_monitor_width)
"""

class CustomDialog(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.conn = conn
        self.title("Unsaved Changes")

        tk.Label(self, text="You have unsaved changes.").pack(padx=10, pady=10)

        self.save_button = tk.Button(self, text="Save", command=self.on_save)
        self.save_button.pack(side=tk.RIGHT, padx=(10, 10), pady=(10, 10))

        self.cancel_button = tk.Button(self, text="Cancel", command=self.on_cancel)
        self.cancel_button.pack(side=tk.RIGHT, padx=(10, 10), pady=(10, 10))

        self.delete_button = tk.Button(self, text="Delete", command=self.on_delete)
        self.delete_button.pack(side=tk.LEFT, padx=(10, 10), pady=(10, 10))

        # Center the dialog on the screen
        self.geometry("+{}+{}".format(parent.winfo_rootx() + 50, parent.winfo_rooty() + 50))

        # Make this dialog modal
        self.transient(parent)
        self.grab_set()
        parent.wait_window(self)

    def on_save(self):
        export_all_blocks(self.conn, True) 
        global is_saved
        is_saved = True
        self.master.destroy()

    def on_cancel(self):
        self.destroy()

    def on_delete(self):
        if messagebox.askyesno("Confirm Delete", "Unsaved Changes. Are you sure you want to delete?"):
            print("Application terminated")
            self.master.destroy()


def socket_server(port, callback):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(('', port))
        s.listen()
        while True:
            conn, _ = s.accept()
            with conn:
                data = conn.recv(1024).decode()
                callback(data)

global_sheet = None
global_current_block_view = 1

def clear_sign_records(conn):
    cursor = conn.cursor()
    cursor.execute("DELETE FROM sign_records")
    conn.commit()

def convert_excel_to_sqlite(excel_file_path):
    df = pd.read_excel(excel_file_path)
    conn = sqlite3.connect('students.db')
    df.to_sql('students', conn, if_exists='replace', index=False)
    return conn

def clear_display(text_widget):
    text_widget.delete('1.0', tk.END)

def record_sign_in_out(conn, rfid_id):
    current_time = datetime.datetime.now().strftime("%H:%M:%S")
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM sign_records WHERE rfid_id = ? ORDER BY id DESC LIMIT 1", (rfid_id,))
    last_record = cursor.fetchone()
    
    if last_record:
        if last_record[3] and not last_record[4]:  # if there is a sign_in_time but no sign_out_time
            cursor.execute("UPDATE sign_records SET sign_out_time = ? WHERE id = ?", (current_time, last_record[0]))
        else:
            cursor.execute("INSERT INTO sign_records (student_id, rfid_id, sign_in_time) VALUES ((SELECT id FROM students WHERE RFIDID = ?), ?, ?)", (rfid_id, rfid_id, current_time))
    else:
        cursor.execute("INSERT INTO sign_records (student_id, rfid_id, sign_in_time) VALUES ((SELECT id FROM students WHERE RFIDID = ?), ?, ?)", (rfid_id, rfid_id, current_time))
    conn.commit()

def export_all_blocks(conn, killApp):
    global is_saved
    # Ask user where to save the file
    filepath = filedialog.asksaveasfilename(
        defaultextension=".xlsx", 
        filetypes=[("Excel files", "*.xlsx")], 
        initialfile=datetime.datetime.now().strftime("%m-%d-%Y.xlsx")
    )

    if filepath:
        with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
            # Iterate over each block and write its data to a separate sheet
            for block_number in range(1, 5):
                df_block = create_records_dataframe(conn, block_number)
                # Ensure headers are included
                df_block.to_excel(writer, sheet_name=f'Block {block_number}', index=False)
                is_saved = True
        if killApp:
            root.quit()
        pass

def auth_and_upload_threaded():
    # Run the auth_and_upload function in a new thread
    threading.Thread(target=auth_and_upload, daemon=True).start()

def auth_and_upload():
    global authStatus, drive_button
    drive_button.config(text="Exporting...")
    conn_thread = sqlite3.connect('students.db')
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('/Users/lucentlu/Code Projects/LPCsystem/credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    service = build('drive', 'v3', credentials=creds)
    authStatus = True
    # Create an XLSX file in a temporary location
    temp_file = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
    with pd.ExcelWriter(temp_file.name, engine='openpyxl') as writer:
        for block_number in range(1, 5):
            df_block = create_records_dataframe(conn_thread, block_number)
            df_block.to_excel(writer, sheet_name=f'Block {block_number}', index=False)

    # Prepare the file for uploading
    file_metadata = {'name': datetime.datetime.now().strftime("%m-%d-%Y.xlsx")}
    media = MediaFileUpload(temp_file.name, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    # Upload the file
    file = service.files().create(body=file_metadata, media_body=media, fields='id').execute()
    drive_button.config(text="Uploading...")

    drive_button.config(text=f"Student data uploaded for {datetime.datetime.now().strftime("%m-%d-%Y")} ✅")
    drive_button.after(2000, lambda: drive_button.config(text="Export to Google Drive"))
    is_saved = True
    # Clean up the temporary file
    os.remove(temp_file.name)

def logout():
    global authStatus
    if os.path.exists('token.pickle'):
        os.remove('token.pickle')
    authStatus = False
    drive_button.config(text="Login and Export")
    print("Logged out of Google Drive")

def process_card_number(card_id_str, text_widget, text_widget_secondary):
    global is_saved
    try:
        card_id = int(card_id_str)

        with sqlite3.connect('students.db') as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM students WHERE RFIDID = ?", (card_id,))
            student = cursor.fetchone()

            if student:
                is_saved = False
                record_sign_in_out(conn, card_id)
                student_info = pd.read_sql_query("SELECT * FROM students WHERE RFIDID = ?", conn, params=(card_id,))
                cursor.execute("""
                    SELECT sign_in_time, sign_out_time FROM sign_records
                    WHERE rfid_id = ? ORDER BY id DESC LIMIT 1
                """, (card_id,))
                last_record = cursor.fetchone()

                # Determine the current status
                status = "IN" if last_record and last_record[1] else "OUT"
                display_student_info(text_widget, student_info.iloc[0], status)
                display_student_info(text_widget_secondary, student_info.iloc[0], status)
            else:
                clear_display(text_widget)
                text_widget.config(bg="red")
                text_widget.tag_configure('center', justify='center', font=("TkDefaultFont", 30))
                text_widget.insert(tk.END, f"No student found with RFIDID: {card_id}", 'center')
                text_widget.after(2000, lambda: clear_display(text_widget))

    except ValueError:
        clear_display(text_widget)
        text_widget.config(bg="red")
        text_widget.tag_configure('center', justify='center', font=("TkDefaultFont", 30))
        text_widget.insert(tk.END, "Invalid input. Please enter a valid card number.", 'center')
        text_widget.after(2000, lambda: clear_display(text_widget))
    except sqlite3.Error as e:
        print(f"SQLite error: {e}")


def display_student_info(text_widget, student, status):
    current_time = datetime.datetime.now()
    time_str = current_time.strftime("%H:%M:%S")  # 24-hour time
    date_str = current_time.strftime("%d/%m/%Y (%a)")  # Date with day of the week

    clear_display(text_widget)

    formatted_info = (
        f"Name: {student['FirstName']} {student['LastName']} ({student['CommonName']})\n"
        f"Year: {student['Year']}\n"
        f"Room: {student['Block']}/{student['RoomNumber']}\n"
        f"Tutor: {student['TutorName']}\n"
        f"Current Time: {time_str}\n"
        f"Current Day: {date_str}\n"
        f"STATUS: {status}"
    )

    # Determine if it's a weekend
    is_weekend = current_time.weekday() in [4, 5]  # 0 is Monday, 6 is Sunday
    current_minutes = current_time.hour * 100 + current_time.minute
    #print(current_minutes)

    # Determine background color
    if not is_weekend:
        #weekday
        if current_minutes <= allowedReturnTimeWeekday and current_minutes <= allowedLeaveTime:
            text_widget.config(bg="black")
        elif (current_minutes-allowedReturnTimeWeekday)<=10 and (allowedLeaveTime-current_minutes)<=10:
            text_widget.config(bg="#FFCC00", fg="black")
            text_widget.after(2000, lambda: text_widget.config(bg="black", fg="white"))
        else:
            text_widget.config(bg="red")
            text_widget.after(2000, lambda: text_widget.config(bg="black", fg="white"))
    else:
        if current_minutes <= allowedReturnTimeWeekday and current_minutes >= allowedReturnTimeWeekend:
            text_widget.config(bg="black")
        elif (current_minutes-allowedReturnTimeWeekday)<=10 and (allowedReturnTimeWeekend-current_minutes)<=10:
            text_widget.config(bg="#FFCC00", fg="black")
            text_widget.after(2000, lambda: text_widget.config(bg="black", fg="white"))
        else:
            text_widget.config(bg="red")
            text_widget.after(2000, lambda: text_widget.config(bg="black", fg="white"))

    update_sheet()  
    text_widget.insert(tk.END, formatted_info, 'center')
    text_widget.tag_configure('center', justify='center', font=("TkDefaultFont", fontSize))
    text_widget.after(2000, lambda: text_widget.config(bg="black"))
    text_widget.after(2000, lambda: clear_display(text_widget))

def update_sheet():
    global global_sheet, global_current_block_view

    try:
        conn = sqlite3.connect('students.db')
        if global_sheet is not None:
            df_block = create_records_dataframe(conn, global_current_block_view)
            global_sheet.set_sheet_data(data=df_block.values.tolist())
            global_sheet.headers(df_block.columns.tolist())
        conn.close()
    except Exception as e:
        print(f"Error in update_sheet: {e}")

def create_sign_records_table(conn):
    # This function creates the sign_records table if it doesn't exist.
    cursor = conn.cursor()
    create_table_query = '''
    CREATE TABLE IF NOT EXISTS sign_records (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        student_id INTEGER NOT NULL,
        rfid_id INTEGER NOT NULL,
        sign_in_time TEXT,
        sign_out_time TEXT,
        FOREIGN KEY(student_id) REFERENCES students(id),
        FOREIGN KEY(rfid_id) REFERENCES students(RFIDID)
    );
    '''
    cursor.execute(create_table_query)
    conn.commit()

def create_records_dataframe(conn, block):
    cursor = conn.cursor()
    # Get the list of all students in the block, regardless of sign-in/out status.
    cursor.execute("""
        SELECT id, CommonName || ' ' || LastName AS Name, Block || '/' || RoomNumber AS Room, TutorName, RFIDID
        FROM students
        WHERE Block = ?
        ORDER BY RoomNumber, LastName, FirstName
    """, (block,))
    students = cursor.fetchall()
    
    # Get the sign-in/out records for the block.
    cursor.execute("""
        SELECT students.RFIDID, sign_in_time, sign_out_time
        FROM sign_records
        JOIN students ON students.RFIDID = sign_records.rfid_id
        WHERE students.Block = ?
    """, (block,))
    sign_records = cursor.fetchall()
    
    # Process the sign records into a dictionary with the RFIDID as the key.
    sign_dict = {}
    for record in sign_records:
        rfid_id = record[0]
        if rfid_id not in sign_dict:
            sign_dict[rfid_id] = []
        sign_dict[rfid_id].extend(record[1:])

    # Determine the maximum number of sign-in/out events for any student.
    max_sign_count = max((len(records) for records in sign_dict.values()), default=0)

    # Create sign_columns based on the maximum number of sign-in/out events.
    sign_columns = []
    for i in range(max_sign_count):
        sign_columns.append('Sign in' if i % 2 == 1 else 'Sign out')

    # Build the DataFrame row by row, including all students.
    data = []
    columns = ['Name', 'Room', 'Tutor']
    columns.extend(sign_columns)

    for student in students:
        rfid_id = student[4]
        sign_times = sign_dict.get(rfid_id, [])
        sign_times.extend([None] * (max_sign_count - len(sign_times))) # Ensure correct number of sign in/out columns.
        row = list(student[1:4]) + sign_times
        data.append(row)

    df_records = pd.DataFrame(data, columns=columns)
    return df_records

def change_block_and_update(block_number):
    global global_current_block_view
    global_current_block_view = block_number
    update_sheet()  # Refresh the sheet with the new block data

def view_records(conn):
    global global_sheet
    global global_current_block_view

    # Create a new window
    records_window = Toplevel()
    records_window.title("Record Viewer")
    records_window.geometry("800x600")

    # Create a Sheet widget
    sheet = tksheet.Sheet(records_window)
    sheet.pack(expand=True, fill='both')
    global_sheet = sheet

    # Add buttons for each block
    for block_number in range(1, 5):
        button = tk.Button(records_window, text=f'Block {block_number}',
                           command=lambda bn=block_number: change_block_and_update(bn))
        button.pack(side=tk.LEFT, padx=10, pady=10)

    # Initially populate the sheet with Block 1 records
    change_block_and_update(1)

def confirm_exit(root):
    if is_saved == False:
        CustomDialog(root)
    else:
        root.destroy()

def main():
    global root, second_screen, drive_button
    global is_saved
    global conn
    is_saved = True
    # Database setup
    conn = convert_excel_to_sqlite(student_xlsx_path)
    create_sign_records_table(conn)
    clear_sign_records(conn)

    # Start the socket server in a background thread
    threading.Thread(target=socket_server, args=(int(port), lambda card_number: process_card_number(card_number, display_text, display_text_second_screen)), daemon=True).start()

    if guard_mon_id == student_mon_id-1:
        # Tkinter GUI setup
        root = tk.Tk()
        root.title("School System - Guard Screen")
        root.geometry("1200x800")

        # Tkinter GUI setup for second screen
        second_screen = tk.Toplevel(root)
        second_screen.title("School System - Student Screen")
        second_screen.geometry(f"{student_monitor_width}x{student_monitor_height}+{guard_monitor_width}+0")
        second_screen.attributes('-fullscreen', True)  # This makes the window full screen
    else:
        # Tkinter GUI setup
        root = tk.Tk()
        root.title("School System - Guard Screen")
        root.geometry(f"1200x800+{student_monitor_width}+0")

        # Tkinter GUI setup for second screen
        second_screen = tk.Toplevel(root)
        second_screen.title("School System - Student Screen")
        second_screen.geometry(f"{student_monitor_width}x{student_monitor_height}")
        second_screen.attributes('-fullscreen', True)


    display_text = tk.Text(root, height=35, width=50)
    display_text.pack(expand=True, fill='both')

    export_button = tk.Button(root, text="Export", command=lambda: export_all_blocks(conn, False))
    export_button.pack(side='left', padx=(10, 0))

    drive_button = tk.Button(root, text="Export to Google Drive", command=auth_and_upload_threaded)
    drive_button.pack(side='left', padx=(10, 0))

    logout_button = tk.Button(root, text="Logout", command=logout)
    logout_button.pack(side='left', padx=(10,0))

    view_button = Button(root, text="View", command=lambda: view_records(conn))
    view_button.pack(side='right', padx=(0, 10))

    display_text_second_screen = tk.Text(second_screen, height=35, width=50)
    display_text_second_screen.pack(expand=True, fill='both')

    # Load the logo image
    logo_photo = ImageTk.PhotoImage(Image.open("/Users/lucentlu/Code Projects/LPCsystem/logo2.png"))
    
    # Create a label to display the logo
    logo_label = tk.Label(root, image=logo_photo)
    logo_label.image = logo_photo  # Keep a reference to avoid garbage collection
    logo_label.pack(side=tk.TOP, pady=10)

    # Create a label to display the logo on the second screen
    logo_label_second_screen = tk.Label(second_screen, image=logo_photo)
    logo_label_second_screen.image = logo_photo  # Keep a reference to avoid garbage collection
    logo_label_second_screen.pack(side=tk.TOP, pady=10)

    

    root.protocol("WM_DELETE_WINDOW", lambda: confirm_exit(root))
    
    root.mainloop()
    conn.close()
    print("Application closed.")
    

if __name__ == "__main__":
    main()
    
