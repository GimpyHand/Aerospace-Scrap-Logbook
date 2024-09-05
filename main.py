import tkinter as tk
from file_handler import copy_xlsm_files
from data_handler import load_scrap_parts
from gui_handler import setup_gui
from utilities import open_mailto

# Define the paths
source_dir = r'X:\ENGINE SERVICES\02 Work Orders'
destination_dir = r'X:\ENGINE SERVICES\Scrap Log Files'
scrap_parts_file = 'Scrap Parts Numbers + Descriptions.csv'
csv_file_path = r'X:\AEROSPACE\Aerospace YWG Scrap Parts Logbook\scrap_logbook.csv'

# Copy .xlsm files from source to destination directory
copy_xlsm_files(source_dir, destination_dir)

# Load scrap parts data
scrap_parts = load_scrap_parts(scrap_parts_file)

# Setup the GUI
root = tk.Tk()
root.title("Cadorath Aerospace Scrap Logbook")
root.geometry('1360x768')

# Create frames
form_frame = tk.Frame(root)
form_frame.grid(row=0, column=0, sticky="nsew")

table_frame = tk.Frame(root)
table_frame.grid(row=1, column=0, sticky="nsew")

# Call setup function for the GUI
setup_gui(root, form_frame, table_frame, csv_file_path, scrap_parts)

# Create a status bar with a contact support link
status_frame = tk.Frame(root, bd=1, relief="sunken")
status_frame.grid(row=2, column=0, sticky="ew")
lbl_email_link = tk.Label(status_frame, text="Contact Support", fg="blue", cursor="hand2", anchor='e', padx=5)
lbl_email_link.pack(side="right", fill="x", expand=True)
lbl_email_link.bind("<Button-1>", lambda e: open_mailto())

# Start the Tkinter main loop
root.mainloop()
