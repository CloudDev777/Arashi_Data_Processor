# Authors: Chris Colomb, hellokurisu.io@gmail.com
#          Maika Rabenitas, maikacoleen1205@gmail.com

import csv
import ctypes
import datetime as dt
import os
import shutil

import insert


# returns a list of all .csv files within the directory
def get_csv_files(directory, suffix=".csv"):
    files = os.listdir(directory)
    return [file for file in files if file.endswith(suffix)]


# returns a datetime object of the .csv file
def get_csv_date(csv_file):
    formatted_date = dt.datetime.strptime(csv_file[26:-4], "%Y-%m-%d")
    return formatted_date


# returns the header and rows of the .csv file
def get_csv_data(directory, csv_file):
    file_in_dir = directory + "\\" + csv_file
    file_data = open(file_in_dir)
    reader = csv.reader(file_data)

    next(reader)
    rows = []
    for row in reader:
        rows.append(row)

    file_data.close()

    return rows


# returns the week number (with Monday as first day of the week) from datetime object
def get_week_number(date):
    return date.strftime("%W")


# returns the day of the week from datetime object
def get_day_of_the_week(date):
    return date.strftime("%A")


# returns the number of items sold based on the category and data given
def get_items_sold(category, data):
    for line in data:
        if line[0] == category:
            if category == "Alcohol":
                return line[1]
            elif category == "Appetizers":
                return line[1]
            elif category == "Beverage":
                return line[1]
            elif category == "Bowl":
                return line[1]
            elif category == "Combo":
                return line[1]
            elif category == "Desserts":
                return line[1]
            elif category == "Ramen":
                return line[1]
    return 0


# returns the gross sales based on the category and data given
def get_gross_sales(category, data):
    for line in data:
        if line[0] == category:
            if category == "Alcohol":
                return line[2]
            elif category == "Appetizers":
                return line[2]
            elif category == "Beverage":
                return line[2]
            elif category == "Bowl":
                return line[2]
            elif category == "Combo":
                return line[2]
            elif category == "Desserts":
                return line[2]
            elif category == "Ramen":
                return line[2]
    return 0


def getArray(data):
    array = []
    for i in range(7):
        array.append(['0', '$0'])

    for line in data:
        if line[0] == "Alcohol":
            array[0] = [line[1], line[2]]
        elif line[0] == "Appetizers":
            array[1] = [line[1], line[2]]
        elif line[0] == "Beverage":
            array[2] = [line[1], line[2]]
        elif line[0] == "Bowl":
            array[3] = [line[1], line[2]]
        elif line[0] == "Combo":
            array[4] = [line[1], line[2]]
        elif line[0] == "Desserts":
            array[5] = [line[1], line[2]]
        elif line[0] == "Ramen":
            array[6] = [line[1], line[2]]
    return array


# returns the alphabet representation of the number given
def getColumnLetter(start_column_number):
    if 26 < start_column_number:
        first_letter = int(start_column_number / 26)
        second_letter = start_column_number % 26
        alphabet_first = chr(first_letter + 64)
        alphabet_second = chr(second_letter + 64)
        column_letter = alphabet_first + alphabet_second

        return column_letter

    else:
        column_letter = chr(start_column_number + 64)
        return column_letter


# returns the range to input in the update function in the google sheet API
def getRange(date):
    day_of_the_week = get_day_of_the_week(date)
    week_num = int(get_week_number(date))
    if day_of_the_week == "Saturday" or day_of_the_week == "Sunday":
        column_number = ((week_num * 3) - 1) + 3
    else:
        column_number = ((week_num * 3) - 1)

    range_input = day_of_the_week + "s!" + getColumnLetter(column_number) + "4"
    return range_input


if __name__ == '__main__':
    program_dir = os.path.realpath(os.path.dirname(__file__))
    csv_files = get_csv_files(program_dir)
    count = 0
    for f in csv_files:
        csv_date = get_csv_date(f)
        sheet_range = getRange(csv_date)
        sheet_values = getArray(get_csv_data(program_dir, f))
        insert.sheet.values().update(spreadsheetId=insert.SPREADSHEET_ID, range=sheet_range,
                                     valueInputOption="USER_ENTERED", body={"values": sheet_values}).execute()
        destination = program_dir + "\\COPIED\\" + f
        source = program_dir + "\\" + f
        shutil.move(source, destination)
        count += 1

    if count == 0:
        ctypes.windll.user32.MessageBoxW(0, "No CSV files were found.", "Error", 0)
    else:
        message_text = str(count) + " file(s) of CSV data copied to Google Sheets."
        ctypes.windll.user32.MessageBoxW(0, message_text, "Success!", 0)
