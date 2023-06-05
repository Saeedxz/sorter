import pandas as pd
import os
import shutil

# Path to the Excel file
excel_file = 'path/to/your/excel/file.xlsx'

# Path to the directory where the files are currently located
source_directory = 'path/to/source/directory/'

# Read the Excel file into a pandas DataFrame
df = pd.read_excel(excel_file)

# Iterate over each row in the DataFrame, starting from the second row
for index, row in df.iterrows():
    # Get the string from the first column
    directory_name = str(row.iloc[0])

    # Create the directory if it doesn't exist
    target_directory = os.path.join(source_directory, directory_name)
    os.makedirs(target_directory, exist_ok=True)

    # Move the files to the target directory
    for file_name in row.iloc[1:]:
        if isinstance(file_name, str) and os.path.exists(os.path.join(source_directory, file_name)):
            source_path = os.path.join(source_directory, file_name)
            destination_path = os.path.join(target_directory, file_name)
            shutil.move(source_path, destination_path)
            print(f'Moved {file_name} to {directory_name}')

print('File moving completed.')
