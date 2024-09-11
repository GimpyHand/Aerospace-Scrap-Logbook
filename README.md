# Cadorath Aerospace Scrap Logbook

This project is a PyQt5 application for managing a scrap logbook for Cadorath Aerospace. It allows users to input, store, and view records of scrapped parts in an SQLite database.

## Features

- **Form Input**: Users can input details such as date, work order (W/O), part number, part description, serial number, initials, and remarks.
- **Data Validation**: Ensures that required fields (date and initials) are filled and that the date is in the correct format (YYYY-MM-DD).
- **Database Storage**: Records are stored in an SQLite database.
- **Data Display**: Records are displayed in a table with adjustable column widths.
- **Record Count**: Displays the total number of records in the database.
- **Clear Form**: Allows users to clear the form inputs.
- **Insert Current Date**: Provides a button to insert the current date into the date field.

## Usage

1. Ensure the path to the SQLite database file is correctly set in the `db_file_path` variable in `V3.py`.

2. Run the application:
    ```sh
    python V3.py
    ```

## Project Structure

- `V3.py`: Main application file containing the PyQt5 GUI and logic.
- `V3_Backup.py`: Backup of the main application file.
- `requirements.txt`: List of required Python packages.

## Dependencies

- Python 3.x
- PyQt5
- pandas

## Contributing

1. Fork the repository.
2. Create a new branch (`git checkout -b feature-branch`).
3. Commit your changes (`git commit -am 'Add new feature'`).
4. Push to the branch (`git push origin feature-branch`).
5. Create a new Pull Request.

## License

This project is licensed under the MIT License. See the `LICENSE` file for details.

## Contact

For support, please contact [alex.jessup@cadorath.com](mailto:alex.jessup@cadorath.com?subject=Scrap%20Log%20Issue).
