"""
Compound Class Visual Representation
Copyright (C) 2026 Nitish Kapur
GitHub: [github.com/nitish-kapur](https://github.com/nitish-kapur)
Licensed under GNU GPLv3
"""

"""
    1.  Opens a file dialog in the script's directory for the user to select
        an .xlsx file.
    2.  Reads the selected file into a pandas DataFrame using the openpyxl
        engine.
    3.  Generates a timestamped output filename based on the original file's
        name (e.g. 'MyData_20260328_143000.xls').
    4.  Opens a hidden Excel instance via xlwings and creates a new workbook.
    5.  Writes the column headers to the first row and applies bold formatting
        to them.
    6.  Writes all data rows starting from the second row.
    7.  Saves the workbook as a legacy .xls file in the current working
        directory and closes the Excel instance.
"""

import tkinter as tk
from tkinter import filedialog
import pandas as pd
import os
import xlwings as xw  # Importing xlwings for saving as .xls
from datetime import datetime  # Import datetime module for timestamp


def convert_xlsx_to_xls():
    # Get the directory where the Python script is located
    script_dir = os.path.dirname(os.path.abspath(__file__))

    # Set up the Tkinter root window (hidden)
    root = tk.Tk()
    root.withdraw()  # Hide the root window

    # Open a file dialog to select an .xlsx file, starting from the script's directory
    file_path = filedialog.askopenfilename(title="Select an Excel file",
                                           filetypes=[("Excel Files", "*.xlsx")],
                                           initialdir=script_dir)

    if not file_path:
        print("No file selected. Exiting...")
        return

    # Load the selected .xlsx file using pandas
    try:
        df = pd.read_excel(file_path, engine='openpyxl')
    except Exception as e:
        print(f"Error reading the file: {e}")
        return

    # Get the filename without the extension
    file_name = os.path.basename(file_path)
    file_name_without_extension = os.path.splitext(file_name)[0]

    # Generate a timestamp to append to the filename
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    # Define the output path for the .xls file, including timestamp in the name
    output_file_name = f"{file_name_without_extension}_{timestamp}.xls"
    output_path = os.path.join(os.getcwd(), output_file_name)

    # Save the dataframe as an .xls file using xlwings
    try:
        # Create a new workbook and add the dataframe to it
        with xw.App(visible=False) as app:  # Open Excel invisibly
            wb = app.books.add()  # Create a new workbook
            sheet = wb.sheets[0]  # Access the first sheet

            # Write the header (column names) in the first row, starting at A1
            sheet.range('A1').value = df.columns.tolist()  # Column headers

            # Bold the header row
            sheet.range('A1').expand('right').api.Font.Bold = True

            # Write the data starting from the second row
            sheet.range('A2').value = df.values  # Write the values of the dataframe (excluding index)

            # Save the file
            wb.save(output_path)
            print(f"File saved as: {output_path}")
            wb.close()  # Close the workbook
    except Exception as e:
        print(f"Error saving the file: {e}")


if __name__ == "__main__":
    convert_xlsx_to_xls()
