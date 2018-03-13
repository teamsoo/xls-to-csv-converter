import xlrd
import unicodecsv
import os
import sys
import datetime


def run_convert(input_file_name, output_file_name, output_delim=","):
    print("Converting file from {} to csv".format(
        os.path.abspath(input_file_name)
    ))
    file_path = os.path.dirname(os.path.abspath(input_file_name))
    print("Opening input file")
    workbook = xlrd.open_workbook(input_file_name)
    worksheets = workbook.sheets()
    for worksheet in worksheets:
        full_output_file_name = "{}/{}_{}.csv".format(
            file_path,
            output_file_name.replace('.csv', ''),
            worksheet.name
        )
        print("Output file initiated: {}".format(full_output_file_name))
        print("Converting {} to {} ...".format(
            worksheet.name,
            output_file_name
        ))
        csv_file = open(
            full_output_file_name,
            'wb',
        )
        csv_writer = unicodecsv.writer(
            csv_file, delimiter=output_delim, encoding='utf-8')
        num_rows = worksheet.nrows - 1
        num_cells = worksheet.ncols - 1
        curr_row = -1
        while curr_row < num_rows:
            curr_row += 1
            row = worksheet.row(curr_row)
            curr_cell = -1
            csv_row = []
            while curr_cell < num_cells:
                curr_cell += 1
                cell_value = get_cell_data(
                    workbook, worksheet, curr_row, curr_cell)
                cell_value = str(cell_value).strip()
                csv_row.append(cell_value)
            if not is_blank_data(csv_row):
                csv_writer.writerow(csv_row)
        csv_file.close()
        print("Worksheet {} converted, {} rows".format(
            worksheet.name, curr_row + 1
        ))
        print("File is saved as {}".format(full_output_file_name))


def is_blank_data(csv_row):
    for data in csv_row:
        if data:
            return False
    return True


def get_cell_data(workbook, worksheet, row, cell):
    cell_type = worksheet.cell_type(row, cell)
    cell_value = worksheet.cell_value(row, cell)
    cell_value_c = convert_type(
        cell_value, cell_type, workbook.datemode)
    return cell_value_c


def convert_type(cell_value, cell_type, date_mode):
    if cell_type == 3:
        year, month, day, hour, minute, second = xlrd.xldate_as_tuple(
            cell_value, date_mode
        )
        if (year == 0):
            if (len(str(hour)) == 1):
                hour = "0{}".format(hour)
            if (len(str(minute)) == 1):
                minute = "0{}".format(minute)
            if (len(str(second)) == 1):
                second = "0{}".format(second)
            data = "{}{}{}".format(hour, minute, second)
        else:
            new_month = str(month)
            if (len(new_month) == 1):
                new_month = "0{}".format(new_month)
            new_day = str(day)
            if (len(new_day) == 1):
                new_day = "0{}".format(new_day)
            data = "{}{}{}".format(year, new_month, new_day)
    elif cell_type == 2:
        if cell_value % 1 == 0:
            data = int(cell_value)
        else:
            data = cell_value
    else:
        data = cell_value
    return data


if __name__ == "__main__":
    try:
        input_file_name = sys.argv[1]
        output_file_name = sys.argv[2]
        if os.path.isfile(input_file_name):
            try:
                delim = sys.argv[3]
                if len(delim) > 1:
                    print("Delimiter should 1 character")
                else:
                    run_convert(input_file_name, output_file_name, delim)
            except:
                run_convert(input_file_name, output_file_name)
        else:
            print("Input file is not valid")
    except IndexError:
        print(
            """Converter should be called as: \n\
python converter.py <input_file.xlsx> <output_file_name_prefix> [output delimiter] \n\
ex. python converter.py excel.xlsx csv_file "|"
            """
        )

