import tkinter as tk
from screeninfo import get_monitors
import ctypes
import pandas as pd
import customtkinter as ctk
from customtkinter import CTkImage
from tkinter import scrolledtext, ttk, messagebox, Toplevel, Button, filedialog
from pandas import ExcelWriter
import sqlite3, datetime, os, tempfile, sys, subprocess, shutil, tksheet, socket, threading, re , platform, cv2, imageio
from PIL import Image, ImageTk
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import os.path
import pickle
from googleapiclient.http import MediaFileUpload
from winotify import Notification

# User32 DLL
user = ctypes.windll.user32
allowedReturnTimeWeekday = 2100
allowedReturnTimeWeekend = 2300
bufferTime = 10
fontSize = 40

SCOPES = ['https://www.googleapis.com/auth/drive']

sheetsLive = False
authStatus = False
global_sheet = None
global_current_block_view = 1
tempFile = ""

# RECT structure
class RECT(ctypes.Structure):
    _fields_ = [
        ('left', ctypes.c_long),
        ('top', ctypes.c_long),
        ('right', ctypes.c_long),
        ('bottom', ctypes.c_long)
    ]
    
    # Method to return the rectangle as a list
    def dump(self):
        return [int(val) for val in (self.left, self.top, self.right, self.bottom)]

# MONITORINFO structure
class MONITORINFO(ctypes.Structure):
    _fields_ = [
        ('cbSize', ctypes.c_ulong),
        ('rcMonitor', RECT),
        ('rcWork', RECT),
        ('dwFlags', ctypes.c_ulong)
    ]

#<-----start camera----->
# Function to find available cameras
def find_cameras():
    index = 0
    arr = []
    while True:
        cap = cv2.VideoCapture(index, cv2.CAP_DSHOW)
        if not cap.read()[0]:
            break
        else:
            arr.append(index)
        cap.release()
        index += 1
    return arr

def take_photo(camera_index, info):
    if getattr(sys, 'frozen', False):
        app_dir = os.path.dirname(sys.executable)
    else:
        app_dir = os.path.dirname(os.path.abspath(__file__))
    main_dir_path = os.path.join(app_dir, 'Photos')
    
    if not os.path.exists(main_dir_path):
        os.makedirs(main_dir_path)
    
    for i in range(1, 5):
        dir_path = os.path.join(main_dir_path, f'Block-{i}')
        if not os.path.exists(dir_path):
            os.makedirs(dir_path)

    try: 
        reader = imageio.get_reader(f'<video{camera_index}>')
        print(f'<video{camera_index}>')
        img = reader.get_next_data()
    except Exception as e:
        print(f"Error in taking photo: {e}")
        return

    # Generate a timestamp for the filename
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    filename = f"{info['CommonName']}_{info['Block']}-{info['RoomNumber']}_{timestamp}.jpg"
    # Save the photo in the 'photos' directory
    block_folder = f"Block-{info['Block']}"
    pic_path = os.path.join(main_dir_path, block_folder)
    imageio.imwrite(os.path.join(pic_path, filename), img)
    print(f"Photo taken and saved as {filename} at {pic_path}")

#<-------end camera-------->

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

# Function to get all monitors
def get_monitors():
    retval = []
    CBFUNC = ctypes.WINFUNCTYPE(ctypes.c_int, ctypes.c_ulong, ctypes.c_ulong, ctypes.POINTER(RECT), ctypes.c_double)
    
    # Callback function for EnumDisplayMonitors
    def cb(hMonitor, hdcMonitor, lprcMonitor, dwData):
        r = lprcMonitor.contents
        data = [hMonitor]
        data.append(r.dump())
        retval.append(data)
        return 1
    
    cbfunc = CBFUNC(cb)
    user.EnumDisplayMonitors(0, 0, cbfunc, 0)
    return retval

# Function to get the monitor areas
def monitor_areas():
    retval = []
    monitors = get_monitors()
    for hMonitor, extents in monitors:
        mi = MONITORINFO()
        mi.cbSize = ctypes.sizeof(MONITORINFO)
        mi.rcMonitor = RECT()
        mi.rcWork = RECT()
        user.GetMonitorInfoA(hMonitor, ctypes.byref(mi))
        data = mi.rcMonitor.dump()
        retval.append(data)
    return retval

def validate_port(*args):
    port = port_var.get()
    if port.isdigit():
        port_label.configure(text="Port (Defult: 12345) ✅")
    else:
        port_label.configure(text="Port (Defult: 12345) ❌")

def validate_database(*args):
    database = student_xlsx_path_var.get()
    if database != None:
        student_xlsx_path_label.configure(text="Database ✅")
    else:
        student_xlsx_path_label.configure(text="Database ❌")

def start_application():
    global cameraChoice, drive_path
    port = port_entry.get()
    student_xlsx_path = student_xlsx_path_var.get()
    drive_path = drive_path_var.get()
    guard_screen = monitors[int(guard_var.get().split(" ")[1])] #[Left, Top, Right, Bottom]
    student_screen = monitors[int(student_var.get().split(" ")[1])]
    cameraChoice = camera_var.get()
    # Validate the inputs
    if not port.isdigit():
        error_label.configure(text="Invalid port.")
    elif not student_xlsx_path:
        error_label.configure(text="Invalid student database path.")
    elif not drive_path:
        error_label.configure(text="Invalid drive path.")
    elif guard_screen == student_screen:
        error_label.configure(text="Error: Guard Screen and Student Screen cannot be the same.")
    elif cameraChoice == None:
        error_label.configure(text="Error: Camera not selected.")
    else:
        print("Application started with Guard Screen: {} and Student Screen: {}".format(guard_var.get(), student_var.get()))
        print(f"Port: {port}, Student XLSX Path: {student_xlsx_path}")
        lpcsys(str(guard_screen[0]), str(guard_screen[1]), str(guard_screen[2]), str(guard_screen[3]), str(student_screen[0]), str(student_screen[1]), str(student_screen[2]), str(student_screen[3]), port, student_xlsx_path, drive_path, cameraChoice)
        print("startapp successful")

def browse_file():
    filename = filedialog.askopenfilename(filetypes=[('Excel Files', '*.xlsx')])
    student_xlsx_path_var.set(filename)

#<-------start main-------->
def extract_drive_id(url):
    # This regex pattern assumes that the ID will be a series of characters
    # that come after 'folders/' and end either at a slash or at the end of the URL
    match = re.search(r'/folders/([a-zA-Z0-9_-]+)', url)
    if match:
        # Return the matched group, which is the folder ID
        return match.group(1)
    else:
        return None

class CustomDialog(tk.Toplevel):
    global conn
    def __init__(self, parent):
        super().__init__(parent)
        self.conn = conn
        self.title("Unsaved Changes")
        self.iconbitmap(resource_path("icon.ico"))

        tk.Label(self, text="You have unsaved changes.").pack(padx=10, pady=10)

        self.save_button = tk.Button(self, text="Save", command=self.on_save)
        self.save_button.pack(side=tk.RIGHT, padx=(10, 10), pady=(10, 10))

        self.cancel_button = tk.Button(self, text="Cancel", command=self.on_cancel)
        self.cancel_button.pack(side=tk.RIGHT, padx=(10, 10), pady=(10, 10))

        self.delete_button = tk.Button(self, text="Delete", command=self.on_delete)
        self.delete_button.pack(side=tk.LEFT, padx=(10, 10), pady=(10, 10))

        # Center the dialog on the screen
        self.geometry(f"+{int((gR-gL)/2)-20}+{int((gB-gT)/2)-20}")

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
    #current_time = datetime.datetime.now().strftime("%H:%M:%S") #Get date/hour/min/sec instead then display/insert different format into the record and header
    current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
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
        # Create a new connection for the thread
        thread_conn = sqlite3.connect('students.db')
        with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
            # Iterate over each block and write its data to a separate sheet
            for block_number in range(1, 5):
                df_block = create_records_dataframe(thread_conn, block_number)
                # Ensure headers are included and at least one sheet is visible
                if not df_block.empty:
                    df_block.to_excel(writer, sheet_name=f'Block {block_number}', index=False)
                else:
                    # If the dataframe is empty, create a dummy sheet to avoid the IndexError
                    writer.book.create_sheet(title=f'Block {block_number}')
                
        is_saved = True
        thread_conn.close()
        if killApp:
            root.quit()
        pass
    
def upload_to_drive(finished_event):
    global authStatus, is_saved, tempFile, service

    toast = Notification(app_id="LPC Gate System",
                     title="File Exported to Drive Successfully",
                     icon=resource_path("icon.ico"))
    toast.add_actions(label="Open Drive",
                 launch=drive_path)
    conn_thread = sqlite3.connect('students.db')
    
    # Create an XLSX file in a temporary location
    temp_file = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
    temp_file_name = temp_file.name
    temp_file.close()
    try: 
        with pd.ExcelWriter(temp_file.name, engine='openpyxl') as writer:
            for block_number in range(1, 5):
                df_block = create_records_dataframe(conn_thread, block_number)
                df_block.to_excel(writer, sheet_name=f'Block {block_number}', index=False)

        # Prepare the file for uploading
        file_metadata = {'name': datetime.datetime.now().strftime("%m-%d-%Y.xlsx"), "parents": [folder_id]} #change to parameters? or location section screen?
        media = MediaFileUpload(temp_file.name, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

        # Upload the file
        try:
            file = service.files().create(body=file_metadata, media_body=media, fields='id').execute()
            print("File uplaoded successfully.")
            toast.show()
        except Exception as e:
            print(f"Error in uploading file: {e}")

    except Exception as e:
        # Handle exceptions that may occur during file creation or processing
        print(f"An error occurred: {e}")

    finally:
        tempFile = temp_file_name
        is_saved = True
        print("is_saved=True")
        finished_event.set()

def open_url(url):
    try: 
        webbrowser.open_new(url)
        print('Opening URL...')  
    except: 
        print('Failed to open URL. Unsupported variable type.')

def upload_threaded():
    finished_event = threading.Event()
    threading.Thread(target=upload_to_drive, args=(finished_event,), daemon=True).start()
    check_thread_finished(finished_event)

def check_thread_finished(finished_event):
    if finished_event.is_set():
        # The thread has finished, perform any follow-up tasks
        on_thread_complete()
    else:
        # If the thread is not finished, check again after a short delay
        root.after(100, check_thread_finished, finished_event)

def on_thread_complete():
    print("Thread has completed its work")
    if os.path.exists(tempFile):
            try:
                os.remove(tempFile)
                print("Temp file removed")
            except PermissionError as e:
                print(f"Could not delete the temporary file: {e}")

def logout():
    global authStatus
    if os.path.exists('token.pickle'):
        os.remove('token.pickle')
    authStatus = False
    print("Logged out of Google Drive")


def process_card_number(card_id_str, text_widget, text_widget_secondary):
    global is_saved, sheetsLive, records_window, conn
    try:
        card_id = int(card_id_str)
        #print("ID length = ", len(card_id_str))

        with sqlite3.connect('students.db') as conn:
            cursor = conn.cursor()
            if len(card_id_str) == 10: 
                cursor.execute("SELECT * FROM students WHERE RFIDID = ?", (card_id,))
            elif len(card_id_str) == 5: 
                cursor.execute("SELECT * FROM students WHERE studentID = ?", (card_id,))
            student = cursor.fetchone()

            #print(student)

            if student:
                is_saved = False
                rfid_id = student[12]
                record_sign_in_out(conn, rfid_id)
                student_info = pd.read_sql_query("SELECT * FROM students WHERE RFIDID = ?", conn, params=(rfid_id,))
                cursor.execute("""
                    SELECT sign_in_time, sign_out_time FROM sign_records
                    WHERE rfid_id = ? ORDER BY id DESC LIMIT 1
                """, (rfid_id,))
                last_record = cursor.fetchone()
                # Determine the current status
                status = "IN" if last_record and last_record[1] else "OUT"
                if sheetsLive == True:
                    view_records_close()
                    display_student_info(text_widget, student_info.iloc[0], status) 
                    display_student_info(text_widget_secondary, student_info.iloc[0], status)
                    take_photo(cameraChoice,student_info.iloc[0])
                else:
                    display_student_info(text_widget, student_info.iloc[0], status)
                    display_student_info(text_widget_secondary, student_info.iloc[0], status)
                    take_photo(cameraChoice,student_info.iloc[0])
            else:
                clear_display(text_widget)
                text_widget.insert(tk.END, f"No student found with RFIDID: {card_id}", 'center')
                text_widget.config(bg="red")
                text_widget.tag_configure('center', justify='center', font=("TkDefaultFont", fontSize))
                text_widget.after(3000, lambda: text_widget.config(bg="black"))
                text_widget.after(3000, lambda: clear_display(text_widget))

    except ValueError:
        clear_display(text_widget)
        text_widget.insert(tk.END, "Invalid input.\nPlease enter a valid card number.", 'center')
        text_widget.config(bg="red")
        text_widget.tag_configure('center', justify='center', font=("TkDefaultFont", fontSize))
        text_widget.after(3000, lambda: text_widget.config(bg="black"))
        text_widget.after(3000, lambda: clear_display(text_widget))
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

    # Determine background color
    if not is_weekend:
        #weekday
        if current_minutes <= allowedReturnTimeWeekday:
            text_widget.config(bg="black")
        elif (current_minutes-allowedReturnTimeWeekday)<=10:
            text_widget.config(bg="#FFCC00", fg="black")
            text_widget.after(3000, lambda: text_widget.config(bg="black", fg="white"))
        else:
            text_widget.config(bg="red")
            text_widget.after(30003000, lambda: text_widget.config(bg="black", fg="white"))
    else:
        if current_minutes <= allowedReturnTimeWeekend:
            text_widget.config(bg="black")
        elif (current_minutes-allowedReturnTimeWeekday)<=10:
            text_widget.config(bg="#FFCC00", fg="black")
            text_widget.after(3000, lambda: text_widget.config(bg="black", fg="white"))
        else:
            text_widget.config(bg="red")
            text_widget.after(3000, lambda: text_widget.config(bg="black", fg="white"))

    update_sheet() 
    text_widget.insert(tk.END, formatted_info, 'center')
    text_widget.tag_configure('center', justify='center', font=("TkDefaultFont", fontSize))
    text_widget.after(3000, lambda: text_widget.config(bg="black"))
    text_widget.after(3000, lambda: clear_display(text_widget))



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
    date_set = []  # To keep track of unique dates for the header
    for record in sign_records:
        rfid_id = record[0]
        sign_in_datetime = pd.to_datetime(record[1])
        sign_out_datetime = pd.to_datetime(record[2]) if record[2] else None
        print(sign_in_datetime)
        print(sign_out_datetime)
        # Add the dates to the date set for the header
        date_set.append(sign_in_datetime.strftime("%m/%d"))
        if sign_out_datetime:
            date_set.append(sign_out_datetime.strftime("%m/%d"))

        # Format times for display in the DataFrame
        sign_in_time = sign_in_datetime.strftime("%H:%M:%S")
        sign_out_time = sign_out_datetime.strftime("%H:%M:%S") if sign_out_datetime else None

        if rfid_id not in sign_dict:
            sign_dict[rfid_id] = []
        sign_dict[rfid_id].extend([sign_in_time, sign_out_time])

    #print(sign_dict)
    # Determine the maximum number of sign-in/out events for any student.
    max_sign_count = max((len(records) for records in sign_dict.values()), default=0)

    # Create sign_columns based on the maximum number of sign-in/out events.
    sign_columns = []
    for i in range(max_sign_count):
        sign_columns.append('Sign in' if i % 2 != 0 else 'Sign out')

    # Build the DataFrame row by row, including all students.
    data = []
    columns = ['Name', 'Room', 'Tutor'] + sign_columns

    for student in students:
        rfid_id = student[4]
        sign_times = sign_dict.get(rfid_id, [])
        sign_times.extend([None] * (max_sign_count - len(sign_times))) # Ensure correct number of sign in/out columns.
        row = list(student[1:4]) + sign_times
        data.append(row)

    df_records = pd.DataFrame(data, columns=columns)

    # Iterate over the DataFrame and add a red "X" for students who have signed out but not signed back in
    for index, row in df_records.iterrows():
        # Check if the last non-null record for the student is a sign-out time
        sign_times = [time for time in row[3:] if pd.notnull(time)]
        if sign_times and len(sign_times) % 2 != 0:  # Odd number of sign times means student is currently signed out
            df_records.at[index, 'Status'] = "❌"
        else:
            df_records.at[index, 'Status'] = "✅"

    # Adjust the DataFrame columns to include the header_dates
    df_records.columns = ['Name', 'Room', 'Tutor'] + sign_columns + ['Status']

    if len(date_set)%2 != 0:
        date_set.append(date_set[-1])

    if len(sign_columns) != 0:
        # Insert the new header list as the first row in the DataFrame
        header_info = [''] * 3 + date_set + ['']
        df_records.loc[-1] = header_info
        df_records.index = df_records.index + 1
        df_records = df_records.sort_index()

    return df_records

def change_block_and_update(block_number):
    global global_current_block_view
    global_current_block_view = block_number
    update_sheet()  # Refresh the sheet with the new block data

def view_records_threaded():
    # Run the auth_and_upload function in a new thread
    threading.Thread(target=view_records, daemon=True).start()

def view_records(conn):
    global global_sheet, global_current_block_view, sheetsLive, records_window

    # Create a new window
    records_window = Toplevel()
    records_window.title("Record Viewer")
    records_window.geometry("800x600")
    sheetsLive = True
    #print(f"Sheets live = {sheetsLive}")

    # Create a Sheet widget
    sheet = tksheet.Sheet(records_window)
    sheet.pack(expand=True, fill='both')
    global_sheet = sheet

    # Add buttons for each block
    for block_number in range(1, 5):
        button = ctk.CTkButton(records_window, text=f'Block {block_number}',
                           command=lambda bn=block_number: change_block_and_update(bn))
        button.pack(side=tk.LEFT, padx=10, pady=10)

    # Initially populate the sheet with Block 1 records
    change_block_and_update(1)
    records_window.protocol("WM_DELETE_WINDOW", view_records_close)

def view_records_close():
    global sheetsLive, records_window
    sheetsLive=False
    #print(f"Sheets live = {sheetsLive}")
    records_window.destroy()


def confirm_exit(root):
    if is_saved == False:
        CustomDialog(root)
    else:
        root.destroy()

class ExportDialog(ctk.CTkToplevel):
    global authStatus
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Export Options")
        self.geometry("300x200")
        self.transient(parent)
        self.grab_set()
        self.local_var = tk.IntVar()
        self.drive_var = tk.IntVar()
        self.after(201, lambda: self.iconbitmap(resource_path("icon.ico")))

        ctk.CTkCheckBox(self, text="Save Locally", variable=self.local_var).pack(pady=10)
        ctk.CTkCheckBox(self, text="Upload to Drive", variable=self.drive_var).pack(pady=10)

        ctk.CTkButton(self, text="Export", command=self.export).pack(pady=20)

    def export(self):
        global authStatus
        if self.local_var.get():
            export_all_blocks(conn, False)
        if self.drive_var.get():
            if not authStatus:
                messagebox.showerror("Error", "Not logged in to Google Drive")
            else:
                upload_threaded()
        self.destroy()

def export_dialog():
    ExportDialog(root)

def login_logout():
    global authStatus, login_logout_button, service
    if authStatus:
        logout()
        login_logout_button.configure(text="Login")
    else:
        try:
            threading.Thread(target=authenticate, daemon=True).start()
        except Exception as e:
            print(f"Error in login: {e}")

def authenticate():
    global authStatus, service, login_logout_button
    try:
        creds = None
        if os.path.exists('token.pickle'):
            with open('token.pickle', 'rb') as token:
                creds = pickle.load(token)
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(resource_path("credentials.json"), SCOPES)
                creds = flow.run_local_server(port=0)
            with open('token.pickle', 'wb') as token:
                pickle.dump(creds, token)
        service = build('drive', 'v3', credentials=creds)
        authStatus = True
        # Update the button text in the main thread
        login_logout_button.after(0, lambda: login_logout_button.configure(text="Logout"))
    except Exception as e:
        print(f"Error in login: {e}")

class ResetDialog(ctk.CTkToplevel):
    global authStatus
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Export Options")
        self.geometry("300x200")
        self.transient(parent)
        self.grab_set()
        self.reset_var = tk.IntVar(value=1)
        self.refresh_var = tk.IntVar()
        self.after(201, lambda: self.iconbitmap(resource_path("icon.ico")))

        ctk.CTkCheckBox(self, text="Reset Student Record", variable=self.reset_var).pack(pady=10)
        ctk.CTkCheckBox(self, text="Refresh Database", variable=self.refresh_var).pack(pady=10)

        ctk.CTkButton(self, text="Enter", command=self.enter).pack(pady=20)

    def enter(self):
        global is_saved
        if self.reset_var.get():
            if not is_saved:
                messagebox.showerror("Error", "Unsaved Changes")
            else:
                # Create a new connection in the current thread
                with sqlite3.connect('students.db') as new_conn:
                    create_sign_records_table(new_conn)
                    clear_sign_records(new_conn)

        if self.refresh_var.get():
            # Create a new connection in the current thread
            with sqlite3.connect('students.db') as new_conn:
                convert_excel_to_sqlite(student_xlsx_path)

        self.destroy()

def reset_record():
    ResetDialog(root)

def lpcsys(guard_left, guard_top, guard_right, guard_bottom, student_left, student_top, student_right, student_bottom, pt, sxp, dp, camChoice):
    global gL, gT, gR, gB, sL, sT, sR, sB, student_xlsx_path, drive_path, port, folder_id
    global root, second_screen, login_logout_button
    global is_saved
    global conn
    driver.destroy()
    print("lpcsys started")
    print("photo taken")
    gL, gT, gR, gB = int(guard_left), int(guard_top), int(guard_right), int(guard_bottom)
    sL, sT, sR, sB = int(student_left), int(student_top), int(student_right), int(student_bottom)
    port = pt
    student_xlsx_path = sxp
    drive_path = dp
    folder_id = extract_drive_id(drive_path)
    is_saved = True
    # Database setup
    conn = convert_excel_to_sqlite(student_xlsx_path)
    create_sign_records_table(conn)
    clear_sign_records(conn)

    # Start the socket server in a background thread
    threading.Thread(target=socket_server, args=(int(port), lambda card_number: process_card_number(card_number, display_text, display_text_second_screen)), daemon=True).start()

    # Tkinter GUI setup
    root = ctk.CTk()
    ctk.set_appearance_mode("dark")
    root.title("School System - Guard Screen")
    root.geometry(f"1100x668+{gL+100}+{gT+50}")
    root.iconbitmap(resource_path("icon.ico"))

    # Tkinter GUI setup for second screen
    second_screen = ctk.CTkToplevel(root)
    second_screen.title("School System - Student Screen")
    second_screen.geometry(f"{300}x{300}+{sL}+0") 
    second_screen.overrideredirect(True)
    second_screen.state('zoomed')

    display_text = tk.Text(root, height=35, width=50)
    display_text.config(bg="black", fg="white")
    display_text.pack(expand=True, fill='both')

    export_button = ctk.CTkButton(root, text="Export", command=export_dialog)
    export_button.pack(side='left', padx=(10, 0))

    login_logout_button = ctk.CTkButton(root, text="Login", command=login_logout)
    login_logout_button.pack(side='left', padx=(10, 0))

    clear_data_button = ctk.CTkButton(root, text="Reset", command=reset_record)
    clear_data_button.pack(side='left', padx=(10, 0))

    view_button = ctk.CTkButton(root, text="View", command=lambda: view_records(conn))
    view_button.pack(side='right', padx=(0, 10))

    display_text_second_screen = tk.Text(second_screen, height=35, width=50)
    display_text_second_screen.config(bg="black", fg="white")
    display_text_second_screen.pack(expand=True, fill='both')

    # Load the logo image
    logo_photo = ctk.CTkImage(light_image=Image.open(resource_path("logo.png")), dark_image=Image.open(resource_path("logo.png")), size=(360,50))
    logo_photo2 = ctk.CTkImage(light_image=Image.open(resource_path("logo.png")), dark_image=Image.open(resource_path("logo.png")), size=(360,50))

    # Create a label to display the logo
    logo_label = ctk.CTkLabel(root, image=logo_photo, text="")
    logo_label.image = logo_photo  # Keep a reference to avoid garbage collection
    logo_label.pack(side=tk.TOP, pady=10)
    # Create a label to display the logo on the second screen
    logo_label_second_screen = ctk.CTkLabel(second_screen, image=logo_photo2, text="")
    logo_label_second_screen.image = logo_photo2  # Keep a reference to avoid garbage collection
    logo_label_second_screen.pack(side=tk.TOP, pady=10)
    
    root.iconbitmap(resource_path("icon.ico"))
    root.protocol("WM_DELETE_WINDOW", lambda: confirm_exit(root))
    conn.close()
    root.mainloop()
    print("Application closed.")

def main():
    global driver, port_var, port_label, port_entry, student_xlsx_path_var, student_xlsx_path_label, student_xlsx_path_entry, browse_button, guard_var, student_var, guard_label, guard_dropdown, student_label, student_dropdown, start_button, error_label, drive_path_label, drive_path_var, drive_path_entry, monitors, camera_var

    driver = ctk.CTk()
    ctk.set_appearance_mode("dark")
    driver.title("LPC Sign in/out System")
    driver.geometry("550x410")  # Set the window size

    # Use CustomTkinter widgets
    error_label = ctk.CTkLabel(driver, text="", text_color="red", wraplength=200)
    error_label.grid(row=10, column=0, columnspan=3)

    driver.attributes('-topmost', True)
    driver.after_idle(driver.attributes, '-topmost', False)

    monitors = monitor_areas()
    monitor_choices = [f"Monitor {index} ({m[2]-m[0]}x{m[3]-m[1]})" for index, m in enumerate(monitors)] #0: Left/1: Top/2: Right/3: Bottom
    cameras_choices = find_cameras()
    cameras_choices = [str(index) for index in cameras_choices]
    print(cameras_choices)

    port_var = tk.StringVar(value='12345')
    port_var.trace_add('write', validate_port)
    student_xlsx_path_var = tk.StringVar()
    student_xlsx_path_var.trace_add('write', validate_database)
    port_label = ctk.CTkLabel(driver, text="Port (Defult: 12345) ✅")
    port_entry = ctk.CTkEntry(driver, textvariable=port_var)
    student_xlsx_path_label = ctk.CTkLabel(driver, text="Student Database Sheet")
    student_xlsx_path_var = tk.StringVar(value=r'C:/Users/Dell/Documents/LPC-guard-system/StudentFakeData.xlsx')
    student_xlsx_path_entry = ctk.CTkEntry(driver, textvariable=student_xlsx_path_var)
    drive_path_label = ctk.CTkLabel(driver, text="Drive Upload Folder URL")
    drive_path_var = tk.StringVar(value=r'https://drive.google.com/drive/u/1/folders/13O2zPO_7Elby7kHtInRBhZ1zxM76jYTt')
    drive_path_entry = ctk.CTkEntry(driver, textvariable=drive_path_var)
    browse_button = ctk.CTkButton(driver, text="Browse", command=browse_file)

    camera_var = tk.IntVar()
    guard_var = tk.StringVar()
    student_var = tk.StringVar()
    student_var.set(monitor_choices[0])
    if len(monitor_choices) > 1:
        guard_var.set(monitor_choices[1])
    else:
        guard_var.set(monitor_choices[0])

    if len(cameras_choices) > 1:
        camera_var.set(cameras_choices[1])

    cam_label = ctk.CTkLabel(driver, text="Camera")
    cam_dropdown = ctk.CTkOptionMenu(driver, variable=camera_var, values=cameras_choices)

    guard_label = ctk.CTkLabel(driver, text="Guard Screen")
    guard_dropdown = ctk.CTkOptionMenu(driver, variable=guard_var, values=monitor_choices)

    student_label = ctk.CTkLabel(driver, text="Student Screen")
    student_dropdown = ctk.CTkOptionMenu(driver, variable=student_var, values=monitor_choices)

    start_button = ctk.CTkButton(driver, text="👉 Start Application", command=start_application)
    start_button.grid(column=2,row=10, padx=(0, 20), pady=(10, 20))

    cam_label.grid(row=0, column=0, padx=(20, 10), pady=(10, 20))
    cam_dropdown.grid(row=0, column=1, padx=(10, 20), pady=(10, 20))
    guard_label.grid(row=1, column=0, padx=(20, 10), pady=(10, 20))
    guard_dropdown.grid(row=1, column=1, padx=(10, 20), pady=(10, 20))
    student_label.grid(row=2, column=0, padx=(20, 10), pady=(10, 20))
    student_dropdown.grid(row=2, column=1, padx=(10, 20), pady=(10, 20))
    port_label.grid(row=5, column=0, padx=(20, 10), pady=(10, 20))
    port_entry.grid(row=5, column=1, padx=(10, 20), pady=(10, 20))
    student_xlsx_path_label.grid(row=6, column=0, padx=(20, 10), pady=(10, 20))
    student_xlsx_path_entry.grid(row=6, column=1, padx=(10, 20), pady=(10, 20))
    browse_button.grid(row=6, column=2, padx=(0, 20), pady=(10, 20))
    drive_path_label.grid(row=7, column=0, padx=(20, 10), pady=(10, 20))
    drive_path_entry.grid(row=7, column=1, padx=(10, 20), pady=(10, 20))

    driver.iconbitmap(resource_path("icon.ico"))
    driver.resizable(False, False) 
    driver.mainloop()

if __name__ == "__main__":
    main()
