import csv
from openpyxl import Workbook


def Convert_Csv_To_Xlsx(csv_file, xlsx_file):
    # Open the CSV file for reading
    with open(csv_file, 'r') as f:
        # Create a CSV reader
        reader = csv.reader(f)

        # Create a new workbook and worksheet
        wb = Workbook()
        ws = wb.active

        # Write the CSV data to the worksheet
        for row in reader:
            ws.append(row)

        # Save the workbook to an XLSX file
        wb.save(xlsx_file)