import sys
import sqlite3
from PyQt5.QtWidgets import (QApplication, QWidget, QMainWindow, QVBoxLayout, QHBoxLayout, QGridLayout, QLabel,
                             QLineEdit, QPushButton, QMessageBox, QComboBox, QFrame, QTableWidget, QTableWidgetItem)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QCursor
import pandas as pd
from datetime import datetime
import webbrowser

# Path to the SQLite database file
db_file_path = r'X:\AEROSPACE\Aerospace YWG Scrap Parts Logbook\scrap_logbook.db'

class ScrapLogbook(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Cadorath Aerospace Scrap Logbook")
        self.setGeometry(280, 156, 1360, 768)

        # Main widget and layout
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.main_layout = QVBoxLayout(self.central_widget)

        # Form layout wrapped in a horizontal box layout to center it
        self.form_layout = QGridLayout()  # Create the form layout
        self.hbox_form_layout = QHBoxLayout()  # Create an outer horizontal layout
        self.hbox_form_layout.addStretch(1)    # Add stretchable space before the form to center it
        self.hbox_form_layout.addLayout(self.form_layout)  # Add the form layout
        self.hbox_form_layout.addStretch(1)    # Add stretchable space after the form to center it
        self.main_layout.addLayout(self.hbox_form_layout)  # Add the horizontal layout to the main layout

        # Table frame
        self.table_frame = QFrame(self)
        self.table_layout = QVBoxLayout(self.table_frame)
        self.main_layout.addWidget(self.table_frame)

        # Status bar frame
        self.status_bar = QFrame(self)
        self.status_layout = QHBoxLayout(self.status_bar)
        self.main_layout.addWidget(self.status_bar)

        # Form Elements
        self.setup_form_elements()

        # Status bar elements
        self.setup_status_bar()

        # Table
        self.setup_table()

        # Initialize database and load records
        self.initialize_db()
        self.load_data()

    def setup_form_elements(self):
        # Date
        self.label_date = QLabel("*Date (YYYY-MM-DD):", self)
        self.label_date.setStyleSheet("color: red;")
        self.form_layout.addWidget(self.label_date, 0, 0)

        self.entry_date = QLineEdit(self)
        self.form_layout.addWidget(self.entry_date, 0, 1)

        self.btn_date = QPushButton("📅", self)
        self.btn_date.clicked.connect(self.insert_current_date)
        self.form_layout.addWidget(self.btn_date, 0, 2)

        # W/O
        self.label_wo = QLabel("W/O:", self)
        self.form_layout.addWidget(self.label_wo, 1, 0)

        self.entry_wo = QLineEdit(self)
        self.form_layout.addWidget(self.entry_wo, 1, 1)

        # Part Number
        self.label_part_number = QLabel("Part Number:", self)
        self.form_layout.addWidget(self.label_part_number, 2, 0)

        self.entry_part_number = QLineEdit(self)
        self.form_layout.addWidget(self.entry_part_number, 2, 1)

        # Part Description
        self.label_part_description = QLabel("Part Description:", self)
        self.form_layout.addWidget(self.label_part_description, 3, 0)

        self.entry_part_description = QLineEdit(self)
        self.form_layout.addWidget(self.entry_part_description, 3, 1)

        # Serial Number
        self.label_serial_number = QLabel("Serial Number:", self)
        self.form_layout.addWidget(self.label_serial_number, 4, 0)

        self.entry_serial_number = QLineEdit(self)
        self.form_layout.addWidget(self.entry_serial_number, 4, 1)

        # Initials
        self.label_initials = QLabel("*Initials:", self)
        self.label_initials.setStyleSheet("color: red;")
        self.form_layout.addWidget(self.label_initials, 5, 0)

        self.entry_initials = QLineEdit(self)
        self.form_layout.addWidget(self.entry_initials, 5, 1)

        # Remarks
        self.label_remarks = QLabel("Remarks:", self)
        self.form_layout.addWidget(self.label_remarks, 6, 0)

        self.entry_remarks = QLineEdit(self)
        self.form_layout.addWidget(self.entry_remarks, 6, 1)

        # Buttons
        self.btn_submit = QPushButton("Submit", self)
        self.btn_submit.clicked.connect(self.submit_data)
        self.form_layout.addWidget(self.btn_submit, 7, 0)

        self.btn_clear = QPushButton("Clear", self)
        self.btn_clear.clicked.connect(self.clear_form)
        self.form_layout.addWidget(self.btn_clear, 7, 1)

    def setup_status_bar(self):
        # Adding the record count label to the left
        self.lbl_record_count = QLabel("Total records: 0", self)
        self.status_layout.addWidget(self.lbl_record_count)

        # Add a spacer to push the "Contact Support" link to the right
        self.status_layout.addStretch(1)

        # Adding the Contact Support link to the right
        self.lbl_email_link = QLabel('<a href="mailto:alex.jessup@cadorath.com?subject=Scrap%20Log%20Issue">Contact Support</a>', self)
        self.lbl_email_link.setOpenExternalLinks(True)
        self.status_layout.addWidget(self.lbl_email_link)

    def setup_table(self):
        self.table_widget = QTableWidget(self)
        self.table_layout.addWidget(self.table_widget)


        # Set minimum column widths (adjust based on your needs)
        self.table_widget.setColumnWidth(0, 150)  # Date
        self.table_widget.setColumnWidth(1, 100)  # W/O
        self.table_widget.setColumnWidth(2, 150)  # Part Number
        self.table_widget.setColumnWidth(3, 200)  # Part Description
        self.table_widget.setColumnWidth(4, 150)  # Serial Number
        self.table_widget.setColumnWidth(5, 100)  # Initials
        self.table_widget.setColumnWidth(6, 200)  # Remarks

    def initialize_db(self):
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

    def submit_data(self):
        date = self.entry_date.text().strip()
        initials = self.entry_initials.text().strip()

        if not date or not initials:
            missing_fields = []
            if not date:
                missing_fields.append("Date")
            if not initials:
                missing_fields.append("Initials")
            QMessageBox.warning(self, "Missing Information", f"Please fill out the following required field(s): {', '.join(missing_fields)}")
            return

        try:
            datetime.strptime(date, '%Y-%m-%d')
        except ValueError:
            QMessageBox.warning(self, "Invalid Date", "Please enter the date in YYYY-MM-DD format.")
            return

        wo = self.entry_wo.text().strip()
        part_number = self.entry_part_number.text().strip()
        part_description = self.entry_part_description.text().strip()
        serial_number = self.entry_serial_number.text().strip()
        remarks = self.entry_remarks.text().strip()

        conn = sqlite3.connect(db_file_path)
        cursor = conn.cursor()
        cursor.execute('''
            INSERT INTO scrap_logbook (date, wo, part_number, part_description, serial_number, initials, remarks)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        ''', (date, wo, part_number, part_description, serial_number, initials, remarks))
        conn.commit()
        conn.close()

        self.load_data()

    def load_data(self):
        conn = sqlite3.connect(db_file_path)
        df = pd.read_sql_query('''
            SELECT date, wo, part_number, part_description, serial_number, initials, remarks
            FROM scrap_logbook ORDER BY date DESC
        ''', conn)
        conn.close()

        self.table_widget.setRowCount(len(df))
        self.table_widget.setColumnCount(len(df.columns))
        self.table_widget.setHorizontalHeaderLabels(df.columns)

        for i, row in df.iterrows():
            for j, value in enumerate(row):
                item = QTableWidgetItem(str(value))
                if j in [0, 1, 2, 4, 5]:  # Indexes of columns to be center justified
                    item.setTextAlignment(Qt.AlignCenter)
                else:
                    item.setTextAlignment(Qt.AlignLeft)
                self.table_widget.setItem(i, j, item)

        # Set column widths after loading data (optional, if not already done in setup_table)
        self.table_widget.setColumnWidth(0, 80)  # Date
        self.table_widget.setColumnWidth(1, 80)  # W/O
        self.table_widget.setColumnWidth(2, 150)  # Part Number
        self.table_widget.setColumnWidth(3, 300)  # Part Description
        self.table_widget.setColumnWidth(4, 120)  # Serial Number
        self.table_widget.setColumnWidth(5, 80)  # Initials
        self.table_widget.setColumnWidth(6, 455)  # Remarks

        # Update the total record count
        self.update_record_count()

    def update_record_count(self):
        conn = sqlite3.connect(db_file_path)
        cursor = conn.cursor()
        cursor.execute('SELECT COUNT(*) FROM scrap_logbook')
        count = cursor.fetchone()[0]
        conn.close()
        self.lbl_record_count.setText(f"Total records: {count}")

    def clear_form(self):
        self.entry_date.clear()
        self.entry_wo.clear()
        self.entry_part_number.clear()
        self.entry_part_description.clear()
        self.entry_serial_number.clear()
        self.entry_initials.clear()
        self.entry_remarks.clear()

    def insert_current_date(self):
        self.entry_date.setText(datetime.now().strftime('%Y-%m-%d'))

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ScrapLogbook()
    window.show()
    sys.exit(app.exec_())
