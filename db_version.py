import tkinter as tk
from tkinter import messagebox, ttk
import pandas as pd
from datetime import datetime
import os
import webbrowser
import sqlite3

# Path to the SQLite database file
db_file_path = r'X:\AEROSPACE\Aerospace YWG Scrap Parts Logbook\scrap_logbook.db'

# Function to initialize the database
def initialize_db():
    conn = sqlite3.connect(db_file_path)
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS scrap_logbook (
            id INTEGER PRIMARY KEY,
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
    # Get values from the form
    date = entry_date.get().strip()
    initials = entry_initials.get().strip()

    # Check if required fields are filled out
    if not date or not initials:
        missing_fields = []
        if not date:
            missing_fields.append("Date")
        if not initials:
            missing_fields.append("Initials")

        # Show error message box indicating which fields are missing
        messagebox.showerror("Missing Information", f"Please fill out the following required field(s): {', '.join(missing_fields)}")
        return

    # Validate the date format
    try:
        datetime.strptime(date, '%Y-%m-%d')
    except ValueError:
        messagebox.showerror("Invalid Date", "Please enter the date in YYYY-MM-DD format.")
        return

    # If both required fields are filled and valid, proceed with submission
    wo = entry_wo.get().strip()
    part_number = entry_part_number.get().strip()
    part_description = entry_part_description.get().strip()
    serial_number = entry_serial_number.get().strip()
    remarks = entry_remarks.get().strip()

    # Insert the data into the database
    conn = sqlite3.connect(db_file_path)
    cursor = conn.cursor()
    cursor.execute('''
        INSERT INTO scrap_logbook (date, wo, part_number, part_description, serial_number, initials, remarks)
        VALUES (?, ?, ?, ?, ?, ?, ?)
    ''', (date, wo, part_number, part_description, serial_number, initials, remarks))
    conn.commit()
    conn.close()

    # Refresh the displayed data
    load_data()

    # Clear the form after submission without prompt
    entry_date.delete(0, tk.END)
    entry_wo.delete(0, tk.END)
    entry_part_number.delete(0, tk.END)
    entry_part_description.delete(0, tk.END)
    entry_serial_number.delete(0, tk.END)
    entry_initials.delete(0, tk.END)
    entry_remarks.delete(0, tk.END)

# Function to clear the form with confirmation
def clear_form():
    if messagebox.askyesno("Clear Form", "Are you sure you want to clear the form?"):
        entry_date.delete(0, tk.END)
        entry_wo.delete(0, tk.END)
        entry_part_number.delete(0, tk.END)
        entry_part_description.delete(0, tk.END)
        entry_serial_number.delete(0, tk.END)
        entry_initials.delete(0, tk.END)
        entry_remarks.delete(0, tk.END)

# Function to load and display records from the database
def load_data(records_to_load=50):
    print(f"Loading data... {records_to_load} records")
    conn = sqlite3.connect(db_file_path)

    # Query to get the total number of records
    cursor = conn.cursor()
    cursor.execute("SELECT COUNT(*) FROM scrap_logbook")
    total_records = cursor.fetchone()[0]

    # Determine how many records to load
    limit_query = ""
    if records_to_load != 'ALL':
        limit_query = f"LIMIT {records_to_load}"

    # Query to get the records
    df = pd.read_sql_query(
        f"SELECT date, wo, part_number, part_description, serial_number, initials, remarks FROM scrap_logbook ORDER BY date DESC {limit_query}",
        conn)
    conn.close()
    print(f"Data loaded: {len(df)} rows")

    # Clear the current table display
    for widget in table_frame.winfo_children():
        widget.destroy()

    # Define column widths (modify these values as needed)
    column_widths = {
        'date': 120,
        'wo': 80,
        'part_number': 150,
        'part_description': 300,
        'serial_number': 120,
        'initials': 80,
        'remarks': 475
    }

    # Define custom header names
    custom_headers = ['Date', 'W/O', 'Part Number', 'Part Description', 'Serial Number', 'Initials', 'Remarks']

    # Display the table headers with bold text using custom header names
    for i, col in enumerate(custom_headers):
        header = tk.Label(table_frame, text=col, borderwidth=1, relief="solid", bg='lightblue', anchor='center', padx=5, pady=5, font=('Arial', 10, 'bold'))
        header.grid(row=0, column=i, sticky="nsew")
        table_frame.grid_columnconfigure(i, minsize=column_widths.get(df.columns[i], 100))  # Set minimum column width based on data columns

    # Load all data
    for i, row in df.iterrows():
        for j, value in enumerate(row):
            color = 'white' if i % 2 == 0 else 'lightgrey'
            cell = tk.Label(table_frame, text=value, borderwidth=1, relief="solid", bg=color, anchor='center', padx=5,
                            pady=5, wraplength=column_widths[df.columns[j]])
            cell.grid(row=i + 1, column=j, sticky="nsew")

    # Update status bar with the total number of records
    lbl_record_count.config(text=f"Total records: {total_records}")

# Function to insert the current date
def insert_current_date():
    current_date = datetime.now().strftime('%Y-%m-%d')
    entry_date.delete(0, tk.END)
    entry_date.insert(0, current_date)

# Function to handle record selection change
def on_records_change(event):
    selected_value = record_selection.get()
    if selected_value == 'ALL':
        load_data('ALL')
    else:
        load_data(int(selected_value))

# Function to open the mailto link
def open_mailto():
    webbrowser.open('mailto:alex.jessup@cadorath.com?subject=Scrap%20Log%20Issue')

# Setup the GUI window
root = tk.Tk()
root.title("Cadorath Aerospace Scrap Logbook")
root.geometry('1360x768')  # Set the initial window size
root.minsize(1360, 768)  # Set minimum window size to the initial size
root.maxsize(1360, 768)  # Set maximum window size to the initial size

# Create a frame for the form and buttons
form_frame = tk.Frame(root)
form_frame.grid(row=0, column=0, sticky="nsew")

# Create a center frame to hold the form elements
center_frame = tk.Frame(form_frame)
center_frame.grid(row=0, column=0, padx=5, pady=5)

# Create and place form labels and entries with centered text
tk.Label(center_frame, text="*Date (YYYY-MM-DD):", fg="red").grid(row=0, column=0, padx=5, pady=5, sticky='e')
entry_date = tk.Entry(center_frame, justify='center')
entry_date.grid(row=0, column=1, padx=5, pady=5, sticky='ew')

# Create and place the date button
btn_date = tk.Button(center_frame, text="📅", command=insert_current_date)
btn_date.grid(row=0, column=2, padx=5, pady=5, sticky='w')

tk.Label(center_frame, text="W/O:").grid(row=1, column=0, padx=5, pady=5, sticky='e')
entry_wo = tk.Entry(center_frame, justify='center')
entry_wo.grid(row=1, column=1, padx=5, pady=5, sticky='ew')

tk.Label(center_frame, text="Part Number:").grid(row=2, column=0, padx=5, pady=5, sticky='e')
entry_part_number = tk.Entry(center_frame, justify='center')
entry_part_number.grid(row=2, column=1, padx=5, pady=5, sticky='ew')

tk.Label(center_frame, text="Part Description:").grid(row=3, column=0, padx=5, pady=5, sticky='e')
entry_part_description = tk.Entry(center_frame, justify='center')
entry_part_description.grid(row=3, column=1, padx=5, pady=5, sticky='ew')

tk.Label(center_frame, text="Serial Number:").grid(row=4, column=0, padx=5, pady=5, sticky='e')
entry_serial_number = tk.Entry(center_frame, justify='center')
entry_serial_number.grid(row=4, column=1, padx=5, pady=5, sticky='ew')

tk.Label(center_frame, text="*Initials:", fg="red").grid(row=5, column=0, padx=5, pady=5, sticky='e')
entry_initials = tk.Entry(center_frame, justify='center')
entry_initials.grid(row=5, column=1, padx=5, pady=5, sticky='ew')

tk.Label(center_frame, text="Remarks:").grid(row=6, column=0, padx=5, pady=5, sticky='e')
entry_remarks = tk.Entry(center_frame, justify='center')
entry_remarks.grid(row=6, column=1, padx=5, pady=5, sticky='ew')

# Create and place the submit and clear buttons
submit_button = tk.Button(form_frame, text="Submit", command=submit_data)
submit_button.grid(row=1, column=0, padx=10, pady=10, sticky='w')

clear_button = tk.Button(form_frame, text="Clear Form", command=clear_form)
clear_button.grid(row=1, column=0, padx=10, pady=10, sticky='e')

# Create the table frame and grid it
table_frame = tk.Frame(root)
table_frame.grid(row=1, column=0, sticky="nsew")

# Create the status bar at the bottom of the window
status_frame = tk.Frame(root, relief=tk.SUNKEN, borderwidth=1)
status_frame.grid(row=2, column=0, sticky="nsew")

lbl_record_count = tk.Label(status_frame, text="Total records: 0", anchor="w")
lbl_record_count.pack(side=tk.LEFT, padx=10)

mailto_button = tk.Button(status_frame, text="Report an Issue", command=open_mailto)
mailto_button.pack(side=tk.RIGHT, padx=10)

# Record selection options
record_selection = ttk.Combobox(status_frame, values=[50, 100, 200, 'ALL'], state='readonly', width=10)
record_selection.current(0)
record_selection.bind("<<ComboboxSelected>>", on_records_change)
record_selection.pack(side=tk.RIGHT, padx=10)

# Initialize the database and load the initial data
initialize_db()
load_data()

# Start the Tkinter main loop
root.mainloop()
