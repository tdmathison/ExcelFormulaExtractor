import sys
import re
from os import path

regexp_cell_pos = re.compile(r'\$[A-Z]+\$[0-9]+')
excel_data = {}


# Checks if the string concatenates other cell data and merges the first occurrence in
def reconstruct(data):
    result = re.search(regexp_cell_pos, data)
    if result:
        data_modified = data[0: result.start()]

        if data[result.start(): result.end()] in excel_data.keys():
            # merge in the content from the referenced cell
            data_modified += excel_data[data[result.start(): result.end()]]
        else:
            # blank cell that is used to store values at runtime
            data_modified += '(var-cell)'

        data_modified += data[result.end() + 1:]

        return repair(data_modified)

    return None


# Helps strip away some inline concatenation
def repair(data):
    return data.replace('""', '"').replace('"&"', '')


# Function to convert given number into an excel column
def get_col_name(n):
    # initialize output String as empty
    res = ""

    while n > 0:
        index = (n - 1) % 26
        res += chr(index + ord('A'))
        n = (n - 1) // 26

    return res[::-1]


# Locates all excel sheet cells with data and saves the cell location as a key and the content as the value
def dump_excel_sheet_values(filepath):
    with open(filepath) as file:
        lines = file.readlines()
        for row, row_values in enumerate(lines):
            values = row_values.split(',')
            for col, col_value in enumerate(values):
                value_to_save = col_value.strip().replace('\n', '\\n')
                if value_to_save:
                    excel_data[str.format("${}${}", get_col_name(col + 1), row + 1)] = value_to_save


if __name__ == '__main__':
    if not len(sys.argv) == 2:
        print("{} <path-to-csv>".format(sys.argv[0]))
        exit(1)

    if not path.exists(sys.argv[1]):
        print("Passed in file does not appear to exist.")
        exit(2)

    dump_excel_sheet_values(sys.argv[1])
    final_strings = []

    # Run through each key and attempt to reconstruct
    for key in excel_data:
        print("Initial: {}".format(excel_data[key]))

        new_val = reconstruct(excel_data[key])
        new_string = new_val

        while new_val:
            print(new_val)
            new_string = new_val
            new_val = reconstruct(new_string)

        if (new_string is not None) and (new_string != '(var-cell)'):
            final_strings.append(new_string)

        print("Final: {}".format(new_string))
        print("********************************************************")

    print("\nFinal strings list:")
    print("========================================================")
    for s in final_strings:
        print(s)
