import tkinter as tk
from tkinter import messagebox
from data_handler import load_logbook_data, save_to_csv
from utilities import insert_current_date


def setup_gui(root, form_frame, table_frame, csv_file_path, part_lookup):
    """
    Sets up the Tkinter GUI layout, form, and table.
    """

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
            messagebox.showerror("Missing Information",
                                 f"Please fill out the following required field(s): {', '.join(missing_fields)}")
            return

        # Validate the date format
        try:
            from datetime import datetime
            datetime.strptime(date, '%Y-%m-%d')
        except ValueError:
            messagebox.showerror("Invalid Date", "Please enter the date in YYYY-MM-DD format.")
            return

        # If both required fields are filled and valid, proceed with submission
        wo = entry_wo.get().strip()
        part_number = entry_part_number.get().strip()
        part_description = part_lookup.get(part_number, "")  # Get part description from lookup
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
        save_to_csv(csv_file_path, new_data)

        # Refresh the displayed data
        load_data()

        # Clear the form after submission
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

    # Function to load and display the last 20 records from the CSV file
    def load_data():
        df = load_logbook_data(csv_file_path)

        # Clear the current table display
        for widget in table_frame.winfo_children():
            widget.destroy()

        if not df.empty:
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
                header = tk.Label(table_frame, text=col, borderwidth=1, relief="solid", bg='lightblue', anchor='center',
                                  padx=5, pady=5, font=('Arial', 10, 'bold'))
                header.grid(row=0, column=i, sticky="nsew")
                table_frame.grid_columnconfigure(i, minsize=column_widths.get(col, 100))

            # Display the table data
            for i, row in df.iterrows():
                for j, value in enumerate(row):
                    color = 'white' if i % 2 == 0 else 'lightgrey'
                    cell = tk.Label(table_frame, text=value, borderwidth=1, relief="solid", bg=color, anchor='center',
                                    padx=5, pady=5)
                    cell.grid(row=i + 1, column=j, sticky="nsew")

        table_frame.grid_rowconfigure(0, weight=1)

    # Create and place form labels and entries
    tk.Label(form_frame, text="*Date (YYYY-MM-DD):", fg="red").grid(row=0, column=0, padx=5, pady=5, sticky='e')
    entry_date = tk.Entry(form_frame, justify='center')
    entry_date.grid(row=0, column=1, padx=5, pady=5, sticky='ew')

    btn_date = tk.Button(form_frame, text="📅", command=lambda: insert_current_date(entry_date))
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

    # Load and display the data in the table
    load_data()
