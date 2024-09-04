import tkinter as tk
from tkinter import messagebox
import pandas as pd
from datetime import datetime
import os
import webbrowser
import shutil

# Specify the source and copy destination directories for the Engine Shop Excel Files
# source_dir = 'path/to/your/source_directory'
# destination_dir = 'path/to/your/destination_directory'

# Create the destination directory if it doesn't exist
os.makedirs(destination_dir, exist_ok=True)

# Loop through all files in the source directory
for file_name in os.listdir(source_dir):
    # Check if the file is an xlsm file
    if file_name.endswith('.xlsm'):
        # Construct the full file path
        source_file_path = os.path.join(source_dir, file_name)
        destination_file_path = os.path.join(destination_dir, file_name)

        # Copy the file to the destination directory
        shutil.copy2(source_file_path, destination_file_path)
        print(f"Copied: {file_name}")

# Get the path to the current directory
current_dir = os.path.dirname(__file__)

# Load the scrap parts numbers + descriptions file
scrap_parts_file = os.path.join(current_dir, 'Scrap Parts Numbers + Descriptions.csv')
scrap_parts_df = pd.read_csv(scrap_parts_file)

# Rename the columns to 'Part Number' and 'Description'
scrap_parts_df = scrap_parts_df.rename(columns={'Unnamed: 0': 'Part Number', 'Unnamed: 1': 'Description'})

# Create a dictionary for easy lookup
part_lookup = scrap_parts_df.set_index('Part Number')['Description'].to_dict()

# Path to the CSV file
csv_file_path = r'X:\AEROSPACE\Aerospace YWG Scrap Parts Logbook\scrap_logbook.csv'

# Function to submit data to the CSV file
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

        # Clear the current table display
        for widget in table_frame.winfo_children():
            widget.destroy()

        # Define column widths (modify these values as needed)
        column_widths = {
            'Date': 120,
            'W/O': 80,
            'Part Number': 150,
            'Part Description': 300,
            'Serial Number': 120,
            'Initials': 80,
            'Remarks': 475
        }

        # Display the table headers with bold text
        for i, col in enumerate(df.columns):
            header = tk.Label(table_frame, text=col, borderwidth=1, relief="solid", bg='lightblue', anchor='center', padx=5, pady=5, font=('Arial', 10, 'bold'))
            header.grid(row=0, column=i, sticky="nsew")
            table_frame.grid_columnconfigure(i, minsize=column_widths.get(col, 100))  # Set minimum column width

        # Display the table data
        for i, row in df.iterrows():
            for j, value in enumerate(row):
                color = 'white' if i % 2 == 0 else 'lightgrey'
                cell = tk.Label(table_frame, text=value, borderwidth=1, relief="solid", bg=color, anchor='center', padx=5, pady=5, wraplength=column_widths[df.columns[j]])
                cell.grid(row=i + 1, column=j, sticky="nsew")

        # Adjust row and column weights to make sure the table expands
        for i in range(len(df.columns)):
            table_frame.grid_columnconfigure(i, weight=1)
        table_frame.grid_rowconfigure(0, weight=1)

        # Update status bar with the number of records
        num_records = len(df)
        lbl_record_count.config(text=f"# of records: {num_records}")

# Function to insert the current date
def insert_current_date():
    current_date = datetime.now().strftime('%Y-%m-%d')
    entry_date.delete(0, tk.END)
    entry_date.insert(0, current_date)

# Function to open the mailto link
def open_mailto():
    webbrowser.open('mailto:alex.jessup@cadorath.com?subject=Scrap%20Log%20Issue')

# Setup the GUI window
root = tk.Tk()
root.title("Cadorath Aerospace Scrap Logbook")
root.geometry('1360x768') # Set the initial window size
root.minsize(1360, 768)  # Set minimum window size to the initial size
root.maxsize(1360, 768)  # Set maximum window size to the initial size

# Create a frame for the form and buttons
form_frame = tk.Frame(root)
form_frame.grid(row=0, column=0, sticky="nsew")

# Create and place form labels and entries with centered text
tk.Label(form_frame, text="*Date (YYYY-MM-DD):", fg="red").grid(row=0, column=0, padx=5, pady=5, sticky='e')
entry_date = tk.Entry(form_frame, justify='center')
entry_date.grid(row=0, column=1, padx=5, pady=5, sticky='ew')

# Create and place the date button
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

# Create a scrollable frame for the table
scroll_canvas = tk.Canvas(root)
scrollbar = tk.Scrollbar(root, orient="vertical", command=scroll_canvas.yview)
scroll_frame = tk.Frame(scroll_canvas)

# Function to configure scrolling
def on_frame_configure(event):
    scroll_canvas.configure(scrollregion=scroll_canvas.bbox("all"))

scroll_frame.bind("<Configure>", on_frame_configure)

# Place the frame inside the canvas and configure scrolling
scroll_canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
scroll_canvas.grid(row=1, column=0, sticky="nsew")
scrollbar.grid(row=1, column=1, sticky="ns")

scroll_canvas.config(yscrollcommand=scrollbar.set)

# Add a frame for the table to the scrollable frame with padding
outer_frame = tk.Frame(scroll_frame, padx=5, pady=5)
outer_frame.pack(fill="both", expand=True, padx=5, pady=5)

table_frame = tk.Frame(outer_frame)
table_frame.pack(fill="both", expand=True)

# Create the status bar frame
status_frame = tk.Frame(root, bd=1, relief="sunken")
status_frame.grid(row=2, column=0, sticky="ew")

# Add status bar widgets
lbl_record_count = tk.Label(status_frame, text="# of records: 0", anchor='w', padx=5)
lbl_record_count.pack(side="left", fill="x", expand=True)

lbl_email_link = tk.Label(status_frame, text="Contact Support", fg="blue", cursor="hand2", anchor='e', padx=5)
lbl_email_link.pack(side="right", fill="x", expand=True)
lbl_email_link.bind("<Button-1>", lambda e: open_mailto())

# Load and display the records
load_data()

# Configure row and column weights for resizing
root.grid_rowconfigure(0, weight=1)  # Form frame row
root.grid_rowconfigure(1, weight=3)  # Table row
root.grid_rowconfigure(2, weight=0)  # Status bar row

root.grid_columnconfigure(0, weight=1)  # Main column

# Start the GUI event loop
root.mainloop()