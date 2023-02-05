import csv
from openpyxl import Workbook


if __name__ == '__main__':

    # Create excel workbook
    wb = Workbook()
    ws = wb.active

    # Get csv file
    # TODO: Fix to get this user to input
    csv_filename = "example.csv"

    # Write csv data to excel files
    # TODO: Check encoding
    with open(csv_filename, encoding='utf-8-sig') as csv_file:

        # Add csv data to a dictionary
        csv_reader = csv.DictReader(csv_file)

        for csv_row in csv_reader:

            # TODO: Only insert specific columns
            # TODO: Get header data
            ws.append(csv_row)

    wb.save(filename='Schedule.xlsx')