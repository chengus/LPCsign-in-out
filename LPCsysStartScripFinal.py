import tkinter as tk
from tkinter import ttk, filedialog
from screeninfo import get_monitors
import subprocess, platform, os

PiStatus = False

def validate_port(*args):
    port = port_var.get()
    if port.isdigit():
        port_label.config(text="Port ✅")
    else:
        port_label.config(text="Port ❌")

def validate_database(*args):
    database = student_xlsx_path_var.get()
    if database != None:
        student_xlsx_path_label.config(text="Database ✅")
    else:
        student_xlsx_path_label.config(text="Database ❌")

def start_application():
    port = port_entry.get()
    student_xlsx_path = student_xlsx_path_var.get()
    guard_screen = guard_var.get()
    student_screen = student_var.get()
    # Validate the inputs
    if not port.isdigit():
        error_label.config(text="Invalid port.")
    elif not student_xlsx_path:
        error_label.config(text="Invalid student database path.")
    elif guard_screen == student_screen:
        error_label.config(text="Error: Guard Screen and Student Screen cannot be the same.")
    else:
        print("Application started with Guard Screen: {} and Student Screen: {}".format(guard_var.get(), student_var.get()))
        print(f"Port: {port}, Student XLSX Path: {student_xlsx_path}")
        subprocess.Popen(["python3", "/Users/lucentlu/Code Projects/LPCsystem/LPCsysFinal.py", guard_var.get(), student_var.get(), port, student_xlsx_path])
        print("startapp successful")
        root.quit()

def browse_file():
    filename = filedialog.askopenfilename(filetypes=[('Excel Files', '*.xlsx')])
    student_xlsx_path_var.set(filename)
    

root = tk.Tk()
root.title("LPC Sign in/out System")
root.geometry("600x350")  # Set the window size

error_label = tk.Label(root, text="", fg="red", wraplength=200)
error_label.grid(row=10, column=0, columnspan=3)

root.attributes('-topmost', True)
if platform.system() == 'Darwin':
    tmpl = 'tell application "System Events" to set frontmost of every process whose unix id is {} to true'
    script = tmpl.format(os.getpid())
    subprocess.check_call(['/usr/bin/osascript', '-e', script])
root.after_idle(root.attributes, '-topmost', False)

monitor_choices = [f"Monitor {index} ({m.width}x{m.height})" for index, m in enumerate(get_monitors())]

port_var = tk.StringVar(value='12345')
port_var.trace_add('write', validate_port)
student_xlsx_path_var = tk.StringVar()
student_xlsx_path_var.trace_add('write', validate_database)
port_label = tk.Label(root, text="Port (Deful: 12345) ✅")
port_entry = tk.Entry(root, textvariable=port_var)
student_xlsx_path_label = tk.Label(root, text="Student Database Sheet")
student_xlsx_path_var = tk.StringVar(value='/Users/lucentlu/Code Projects/LPCsystem/StudentFakeData.xlsx')
student_xlsx_path_entry = tk.Entry(root, textvariable=student_xlsx_path_var)
browse_button = tk.Button(root, text="Browse", command=browse_file)

guard_var = tk.StringVar()
student_var = tk.StringVar()
guard_var.set(monitor_choices[0])
student_var.set(monitor_choices[0])

guard_label = tk.Label(root, text="Guard Screen")
guard_dropdown = tk.OptionMenu(root, guard_var, *monitor_choices)

student_label = tk.Label(root, text="Student Screen")
student_dropdown = tk.OptionMenu(root, student_var, *monitor_choices)

start_button = tk.Button(root, text="👉 Start Application", command=start_application)
start_button.grid(column=2,row=10)

guard_label.grid(row=0, column=0, padx=(20, 10), pady=(20, 10))
guard_dropdown.grid(row=0, column=1, padx=(10, 20), pady=(20, 10))
student_label.grid(row=1, column=0, padx=(20, 10), pady=(10, 20))
student_dropdown.grid(row=1, column=1, padx=(10, 20), pady=(10, 20))
port_label.grid(row=5, column=0, padx=(20, 10), pady=(10, 20))
port_entry.grid(row=5, column=1, padx=(10, 20), pady=(10, 20))
student_xlsx_path_label.grid(row=6, column=0, padx=(20, 10), pady=(10, 20))
student_xlsx_path_entry.grid(row=6, column=1, padx=(10, 20), pady=(10, 20))
browse_button.grid(row=6, column=2, padx=(0, 20), pady=(10, 20))

root.mainloop()
