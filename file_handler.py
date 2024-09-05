import os
import shutil

def copy_xlsm_files(source_dir, destination_dir):
    """
    Copies .xlsm files from the source directory (recursively) to the destination directory.
    """
    # Create the destination directory if it doesn't exist
    os.makedirs(destination_dir, exist_ok=True)

    # Loop through all files in the source directory recursively
    for dirpath, dirnames, filenames in os.walk(source_dir):
        for file_name in filenames:
            # Check if the file is an xlsm file
            if file_name.endswith('.xlsm'):
                # Construct the full file path
                source_file_path = os.path.join(dirpath, file_name)
                destination_file_path = os.path.join(destination_dir, file_name)

                # Copy the file to the destination directory
                shutil.copy2(source_file_path, destination_file_path)
                print(f"Copied: {file_name}")
