import tkinter as tk
from tkinter import messagebox
import pandas as pd
from datetime import datetime
import os
import sqlite3

# Path to the SQLite database
db_file_path = r'X:\AEROSPACE\Aerospace YWG Scrap Parts Logbook\scrap_logbook.db'

# Number of records to load at a time
LOAD_BATCH_SIZE = 50

# Function to initialize the database and table if they don't exist
def init_db():
    conn = sqlite3.connect(db_file_path)
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS scrap_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT NOT NULL,
            wo TEXT,
            part_number TEXT,
            part_description TEXT,
            serial_number TEXT,
            initials TEXT NOT NULL,
            remarks TEXT
        )
    ''')
    conn.commit()
    conn.close()

# Function to submit data to the database
def submit_data():
    date = entry_date.get().strip()
    initials = entry_initials.get().strip()

    if not date or not initials:
        missing_fields = []
        if not date:
            missing_fields.append("Date")
        if not initials:
            missing_fields.append("Initials")
        messagebox.showerror("Missing Information", f"Please fill out the following required field(s): {', '.join(missing_fields)}")
        return

    try:
        datetime.strptime(date, '%Y-%m-%d')
    except ValueError:
        messagebox.showerror("Invalid Date", "Please enter the date in YYYY-MM-DD format.")
        return

    data = {
        "date": date,
        "wo": entry_wo.get().strip(),
        "part_number": entry_part_number.get().strip(),
        "part_description": entry_part_description.get().strip(),
        "serial_number": entry_serial_number.get().strip(),
        "initials": initials,
        "remarks": entry_remarks.get().strip()
    }

    conn = sqlite3.connect(db_file_path)
    cursor = conn.cursor()
    cursor.execute('''
        INSERT INTO scrap_log (date, wo, part_number, part_description, serial_number, initials, remarks)
        VALUES (:date, :wo, :part_number, :part_description, :serial_number, :initials, :remarks)
    ''', data)
    conn.commit()
    conn.close()

    load_data()
    clear_form()

# Function to clear the form fields
def clear_form():
    entry_date.delete(0, tk.END)
    entry_wo.delete(0, tk.END)
    entry_part_number.delete(0, tk.END)
    entry_part_description.delete(0, tk.END)
    entry_serial_number.delete(0, tk.END)
    entry_initials.delete(0, tk.END)
    entry_remarks.delete(0, tk.END)

# Function to load data from the database and display it
def load_data():
    conn = sqlite3.connect(db_file_path)
    df = pd.read_sql_query("SELECT * FROM scrap_log", conn)
    conn.close()

    for widget in table_frame.winfo_children():
        widget.destroy()

    for i, col in enumerate(df.columns[1:]):  # Skip the ID column
        header = tk.Label(table_frame, text=col, borderwidth=1, relief="solid", bg='lightblue', anchor='center', padx=5, pady=5, font=('Arial', 10, 'bold'))
        header.grid(row=0, column=i, sticky="nsew")

    for i, row in df.iterrows():
        for j, value in enumerate(row[1:]):  # Skip the ID column
            color = 'white' if i % 2 == 0 else 'lightgrey'
            cell = tk.Label(table_frame, text=value, borderwidth=1, relief="solid", bg=color, anchor='center', padx=5, pady=5)
            cell.grid(row=i + 1, column=j, sticky="nsew")

    lbl_record_count.config(text=f"# of records: {len(df)}")

# Function to insert the current date into the date field
def insert_current_date():
    current_date = datetime.now().strftime('%Y-%m-%d')
    entry_date.delete(0, tk.END)
    entry_date.insert(0, current_date)

# Setup the GUI window
root = tk.Tk()
root.title("Cadorath Aerospace Scrap Logbook")
root.geometry('1360x768')
root.minsize(1360, 768)
root.maxsize(1360, 768)

# Create a frame for the form and buttons
form_frame = tk.Frame(root)
form_frame.grid(row=0, column=0, sticky="nsew")

# Create and place form labels and entries with centered text
tk.Label(form_frame, text="*Date (YYYY-MM-DD):", fg="red").grid(row=0, column=0, padx=5, pady=5, sticky='e')
entry_date = tk.Entry(form_frame, justify='center')
entry_date.grid(row=0, column=1, padx=5, pady=5, sticky='ew')

btn_date = tk.Button(form_frame, text="📅", command=insert_current_date)
btn_date.grid(row=0, column=2, padx=5, pady=5, sticky='w')

tk.Label(form_frame, text="W/O:").grid(row=1, column=0, padx=5, pady=5, sticky='e')
entry_wo = tk.Entry(form_frame, justify='center')
entry_wo.grid(row=1, column=1, padx=5, pady=5, sticky='ew')

tk.Label(form_frame, text="Part Number:").grid(row=2, column=0, padx=5, pady=5, sticky='e')
entry_part_number = tk.Entry(form_frame, justify='center')
entry_part_number.grid(row=2, column=1, padx=5, pady=5, sticky='ew')

tk.Label(form_frame, text="Part Description:").grid(row=3, column=0, padx=5, pady=5, sticky='e')
entry_part_description = tk.Entry(form_frame, justify='center')
entry_part_description.grid(row=3, column=1, padx=5, pady=5, sticky='ew')

tk.Label(form_frame, text="Serial Number:").grid(row=4, column=0, padx=5, pady=5, sticky='e')
entry_serial_number = tk.Entry(form_frame, justify='center')
entry_serial_number.grid(row=4, column=1, padx=5, pady=5, sticky='ew')

tk.Label(form_frame, text="*Initials:", fg="red").grid(row=5, column=0, padx=5, pady=5, sticky='e')
entry_initials = tk.Entry(form_frame, justify='center')
entry_initials.grid(row=5, column=1, padx=5, pady=5, sticky='ew')

tk.Label(form_frame, text="Remarks:").grid(row=6, column=0, padx=5, pady=5, sticky='e')
entry_remarks = tk.Entry(form_frame, justify='center')
entry_remarks.grid(row=6, column=1, padx=5, pady=5, sticky='ew')

# Create and place buttons
btn_submit = tk.Button(form_frame, text="Submit", command=submit_data)
btn_submit.grid(row=7, column=0, padx=5, pady=5, sticky='ew')

btn_clear = tk.Button(form_frame, text="Clear", command=clear_form)
btn_clear.grid(row=7, column=1, padx=5, pady=5, sticky='ew')

# Create a frame for the table
table_frame = tk.Frame(root)
table_frame.grid(row=1, column=0, columnspan=2, sticky="nsew")

# Status bar
lbl_record_count = tk.Label(root, text="0 records", anchor='e', bg='lightgrey')
lbl_record_count.grid(row=2, column=0, columnspan=3, sticky='ew')

# Initialize the database and load data
init_db()
load_data()

root.mainloop()
