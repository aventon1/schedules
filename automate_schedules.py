import csv
from openpyxl import Workbook
import re


def write_csv_to_xlsx(csv_filename, ws):
    """
    Write csv file to xlsx file with specified columns and rows

    :param csv_filename: Str of name of csv file
    :param ws: Workbook xlsx object
    :return: ws, provider_name, date_range, practice_name
    """

    # Write csv data to excel files
    # TODO: Check encoding
    with open(csv_filename, encoding='utf-8-sig') as csv_file:

        # Add csv data to a dictionary
        csv_reader = csv.DictReader(csv_file)

        # Variables to get specific header data
        provider_name = ''
        date_range = ''
        practice_name = ''
        iteration = 0
        for csv_row in csv_reader:

            # Capture header data
            if iteration == 0:
                provider_name = get_provider(csv_row['Textbox9'])
                date_range = csv_row['textbox29']
                practice_name = csv_row['PracticeName']
                iteration += 1

            # Insert only these values into rows
            patient_appointment = [csv_row['AppointmentTime'],
                                   csv_row['Patient'],
                                   csv_row['Comments'],
                                   csv_row['PatientEmailAddress'],
                                   csv_row['AppointmentTypeName'],
                                   csv_row['Carrier'],
                                   csv_row['Provider']]
            ws.append(patient_appointment)

    return ws, provider_name, date_range, practice_name


def get_provider(provider_string):
    """
    Helper function to get substring from given string

    :param provider_string: Str that contains provider name
    :return: provider_name
    """

    provider_match = re.search(r'^P.+(,\s)', provider_string)

    if provider_match:
        provider_name = provider_match.group(0)
    else:
        provider_name = provider_string

    return provider_name

# TODO: Formatting and styles
def apply_styles():
    pass

# TODO: print settings
def apply_print_settings():
    pass

if __name__ == '__main__':

    # Create excel workbook
    wb = Workbook()
    ws = wb.active

    # Get csv file
    # TODO: Fix to get this user to input
    csv_filename = "example.csv"

    ws, provider_name, date_range, practice_name = write_csv_to_xlsx(csv_filename, ws)

    # TODO: Formatting and styles
    # TODO: print settings


    wb.save(filename='Schedule.xlsx')