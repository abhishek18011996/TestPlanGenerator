import openpyxl as py

def create_index_dictionary(input_sheet):
    start_counter = 1
    dictionary = {}
    while input_sheet.cell(1, start_counter).value is not None:
        dictionary[input_sheet.cell(1, start_counter).value] = start_counter
        start_counter += 1
    return dictionary