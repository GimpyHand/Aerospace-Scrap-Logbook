import os
import pytest
import sqlite3
import datetime
from PyQt5.QtWidgets import QApplication
from main import ScrapLogbook, db_file_path

@pytest.fixture(scope="module")
def app():
    app = QApplication([])
    yield app
    app.quit()

@pytest.fixture
def scrap_logbook(app):
    window = ScrapLogbook()
    window.show()
    yield window
    window.close()

def test_initialization(scrap_logbook):
    assert scrap_logbook.windowTitle() == "Cadorath Aerospace Scrap Logbook"
    assert scrap_logbook.geometry().width() == 1360
    assert scrap_logbook.geometry().height() == 768

def test_insert_current_date(scrap_logbook):
    scrap_logbook.insert_current_date()
    assert scrap_logbook.entry_date.text() == datetime.now().strftime('%Y-%m-%d')

def test_submit_data(scrap_logbook):
    scrap_logbook.entry_date.setText("2024-09-13")
    scrap_logbook.entry_initials.setText("AJ")
    scrap_logbook.combo_source.setCurrentText("WPG MRO")
    scrap_logbook.submit_data()

    with sqlite3.connect(db_file_path) as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM scrap_logbook WHERE date LIKE '2024-09-13%'")
        result = cursor.fetchone()
        assert result is not None
        assert result[7] == "AJ"
        assert result[8] == "WPG MRO"

def test_clear_form(scrap_logbook):
    scrap_logbook.entry_date.setText("2024-09-13")
    scrap_logbook.entry_initials.setText("AJ")
    scrap_logbook.clear_form()
    assert scrap_logbook.entry_date.text() == ""
    assert scrap_logbook.entry_initials.text() == ""