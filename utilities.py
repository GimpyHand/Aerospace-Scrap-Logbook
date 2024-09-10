from datetime import datetime
import webbrowser
import tkinter as tk

def insert_current_date(entry_date):
    """
    Inserts the current date into a Tkinter entry widget in YYYY-MM-DD format.
    """
    current_date = datetime.now().strftime('%Y-%m-%d')
    entry_date.delete(0, tk.END)
    entry_date.insert(0, current_date)

def open_mailto():
    """
    Opens a mailto link in the default mail client.
    """
    webbrowser.open('mailto:alex.jessup@cadorath.com?subject=Scrap%20Log%20Issue')
