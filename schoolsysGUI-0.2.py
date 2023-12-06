import sqlite3
import pandas as pd
import tkinter as tk
from tkinter import scrolledtext

def convert_excel_to_sqlite(excel_file_path):
    # Read the Excel file into a Pandas DataFrame
    df = pd.read_excel(excel_file_path)

    # Connect to SQLite database
    conn = sqlite3.connect('students.db')

    # Convert the DataFrame to a SQLite table
    df.to_sql('students', conn, if_exists='replace', index=False)

    # Return the database connection
    return conn

def clear_display(text_widget):
    text_widget.delete('1.0', tk.END)

def on_enter_key(event, text_widget, conn, entry_widget):
    try:
        # Extract RFIDID from the event
        card_id_str = event.widget.get()
        card_id = int(card_id_str)
        # Query the database for the student information based on the RFIDID
        query = "SELECT * FROM students WHERE RFIDID = ?"
        student_info = pd.read_sql_query(query, conn, params=(card_id,))

        # Clear the text widget
        text_widget.delete('1.0', tk.END)

        # Display the student information
        if not student_info.empty:
            student = student_info.iloc[0]

            formatted_info = (
                f"Name: {student['FirstName']} {student['LastName']} ({student['CommonName']})\n"
                f"Year: {student['Year']}\n"
                f"Room: {student['Block']}/{student['RoomNumber']}\n"
                f"Tutor: {student['TutorName']}"
            )
            text_widget.insert(tk.END, formatted_info)
            text_widget.after(5000, lambda: clear_display(text_widget))
        else:
            text_widget.insert(tk.END, f"No student found with RFIDID: {card_id}")
    except ValueError:
        text_widget.delete('1.0', tk.END)
        text_widget.insert(tk.END, "Invalid input. Please enter a valid card number.")

    # Clear the entry widget and refocus it for the next input
    entry_widget.delete(0, tk.END)
    entry_widget.focus_set()



def redirect_focus_to_entry(event, entry_widget):
    entry_widget.focus_set()

def main():
    # Read the Excel file into a Pandas DataFrame
    df = pd.read_excel("/Users/lucentlu/Desktop/LPCsystem/StudentFakeData.xlsx")
    df.columns = df.columns.str.strip()

    # Connect to SQLite database
    conn = sqlite3.connect('students.db')

    # Convert the DataFrame to a SQLite table
    df.to_sql('students', conn, if_exists='replace', index=False)

    # Query the database to confirm data is loaded
    df_loaded = pd.read_sql_query("SELECT * FROM students", conn)

    # Display the DataFrame
    #print(df_loaded)  # Display the first few rows

    # Set up the GUI using tkinter
    root = tk.Tk()
    root.title("School System")
    root.geometry("800x400")  # Set window size

    # Hidden entry widget for capturing RFID input
    entry = tk.Entry(root, width=1)
    entry.place(x=-100, y=-100)  # Position off-screen

    # Bind the Enter key
    root.bind('<Return>', lambda event: on_enter_key(event, display_text, conn, entry))

    # Display text box using scrolledtext for scrolling capability
    display_text = scrolledtext.ScrolledText(root, height=35, width=50, takefocus=0)
    display_text.pack(expand=True, fill='both')
    display_text.bind("<Button-1>", lambda event: redirect_focus_to_entry(event, entry))

    # Immediately set focus to the entry widget
    redirect_focus_to_entry(None, entry)

    root.mainloop()

    # Close the database connection
    conn.close()

if __name__ == "__main__":
    main()
