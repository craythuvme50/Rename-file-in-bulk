#!/usr/bin/python3

import subprocess
import sys
import importlib
import os

# Function to install and import packages
def install_and_import(package):
    try:
        importlib.import_module(package)
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])
    finally:
        globals()[package] = importlib.import_module(package)

# Check and import required modules
install_and_import('openpyxl')
install_and_import('pandas')

def rename(data):
    for index, row in data.iterrows():
        old_name = row['Old_name'].strip()
        new_name = row['New_name'].strip()
        try:
            if os.path.exists(old_name):
                os.rename(old_name, new_name)
                print(f'Successfully renamed {old_name} to {new_name}')
            else:
                print(f'Old file does not exist: {old_name}')
        except Exception as e:
            print(f'Exception occurred while renaming {old_name} to {new_name}: {e}')

if __name__ == "__main__":
    try:
        # Read the Excel file
        data = pandas.read_excel('name.xlsx')

        # Execute the rename function
        rename(data)

        # Log completion of the rename function
        print('Completed renaming files.')
    except Exception as e:
        print(f'Error occurred: {e}')
