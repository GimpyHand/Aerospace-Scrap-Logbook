import os
import pandas as pd


def load_scrap_parts(file_path):
    """
    Loads the Scrap Parts Numbers and Descriptions CSV file into a DataFrame
    and returns a dictionary for easy lookup.
    """
    scrap_parts_df = pd.read_csv(file_path)
    scrap_parts_df = scrap_parts_df.rename(columns={'Unnamed: 0': 'Part Number', 'Unnamed: 1': 'Description'})
    return scrap_parts_df.set_index('Part Number')['Description'].to_dict()


def save_to_csv(csv_file_path, new_data):
    """
    Appends new data to the scrap_logbook.csv file.
    """
    df = pd.DataFrame([new_data])

    if os.path.exists(csv_file_path):
        df.to_csv(csv_file_path, mode='a', header=False, index=False)
    else:
        df.to_csv(csv_file_path, mode='w', header=True, index=False)


def load_logbook_data(csv_file_path):
    """
    Loads the scrap logbook CSV data into a DataFrame.
    """
    if os.path.exists(csv_file_path):
        return pd.read_csv(csv_file_path)
    return pd.DataFrame()  # Return an empty DataFrame if the file doesn't exist
