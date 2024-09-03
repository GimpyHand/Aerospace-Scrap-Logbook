import tkinter as tk
from tkinter import messagebox
import pandas as pd
from datetime import datetime
import os

# Path to the CSV file
csv_file_path = r'X:\AEROSPACE\Aerospace YWG Scrap Parts Logbook\scrap_logbook.csv'


# Function to submit data to the CSV file
def submit_data():
    # Get values from the form
    date = entry_date.get()
    wo = entry_wo.get()
    part_number = entry_part_number.get()
    part_description = entry_part_description.get()
    serial_number = entry_serial_number.get()
    initials = entry_initials.get()
    remarks = entry_remarks.get()

    # Validate the date format
    try:
        datetime.strptime(date, '%Y-%m-%d')
    except ValueError:
        messagebox.showerror("Invalid Date", "Please enter the date in YYYY-MM-DD format.")
        return

    # Append the data to the CSV file
    new_data = {
        "Date": date,
        "W/O": wo,
        "Part Number": part_number,
        "Part Description": part_description,
        "Serial Number": serial_number,
        "Initials": initials,
        "Remarks": remarks
    }
    df = pd.DataFrame([new_data])

    if os.path.exists(csv_file_path):
        df.to_csv(csv_file_path, mode='a', header=False, index=False)
    else:
        df.to_csv(csv_file_path, mode='w', header=True, index=False)

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


# Function to load and display the last 20 records from the CSV file
def load_data():
    if os.path.exists(csv_file_path):
        df = pd.read_csv(csv_file_path)
        df = df.tail(20)  # Get the last 20 records

        # Clear the current table display
        for widget in table_frame.winfo_children():
            widget.destroy()

        # Display the table headers
        for i, col in enumerate(df.columns):
            header = tk.Label(table_frame, text=col, borderwidth=1, relief="solid")
            header.grid(row=0, column=i, sticky="nsew")

        # Display the table data
        for i, row in df.iterrows():
            for j, value in enumerate(row):
                cell = tk.Label(table_frame, text=value, borderwidth=1, relief="solid")
                cell.grid(row=i + 1, column=j, sticky="nsew")


# Setup the GUI window
root = tk.Tk()
root.title("Cadorath Aerospace Scrap Logbook")

# Create and place form labels and entries with right-justified text
tk.Label(root, text="Date (YYYY-MM-DD):", anchor='e').grid(row=0, column=0, padx=5, pady=5, sticky='e')
entry_date = tk.Entry(root, justify='right')
entry_date.grid(row=0, column=1, padx=5, pady=5)

tk.Label(root, text="W/O:", anchor='e').grid(row=1, column=0, padx=5, pady=5, sticky='e')
entry_wo = tk.Entry(root, justify='right')
entry_wo.grid(row=1, column=1, padx=5, pady=5)

tk.Label(root, text="Part Number:", anchor='e').grid(row=2, column=0, padx=5, pady=5, sticky='e')
entry_part_number = tk.Entry(root, justify='right')
entry_part_number.grid(row=2, column=1, padx=5, pady=5)

tk.Label(root, text="Part Description:", anchor='e').grid(row=3, column=0, padx=5, pady=5, sticky='e')
entry_part_description = tk.Entry(root, justify='right')
entry_part_description.grid(row=3, column=1, padx=5, pady=5)

tk.Label(root, text="Serial Number:", anchor='e').grid(row=4, column=0, padx=5, pady=5, sticky='e')
entry_serial_number = tk.Entry(root, justify='right')
entry_serial_number.grid(row=4, column=1, padx=5, pady=5)

tk.Label(root, text="Initials:", anchor='e').grid(row=5, column=0, padx=5, pady=5, sticky='e')
entry_initials = tk.Entry(root, justify='right')
entry_initials.grid(row=5, column=1, padx=5, pady=5)

tk.Label(root, text="Remarks:", anchor='e').grid(row=6, column=0, padx=5, pady=5, sticky='e')
entry_remarks = tk.Entry(root, justify='right')
entry_remarks.grid(row=6, column=1, padx=5, pady=5)

# Create and place buttons
btn_submit = tk.Button(root, text="Submit", command=submit_data)
btn_submit.grid(row=7, column=0, padx=5, pady=5)

btn_clear = tk.Button(root, text="Clear", command=clear_form)
btn_clear.grid(row=7, column=1, padx=5, pady=5)

# Create the table frame to display the CSV data
table_frame = tk.Frame(root)
table_frame.grid(row=8, column=0, columnspan=2, padx=5, pady=5)

# Load and display the last 20 entries
load_data()

# Start the GUI event loop
root.mainloop()
