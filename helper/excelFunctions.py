# ----- Imports ----- #
import xlwings as xw  # <-- Current library used for manipulating Excel sheet


# Sets app visibility
def set_app(is_visible: bool):
    if is_visible:
        return xw.App(visible=True)
    else:
        return xw.App(visible=False)


# Loads the Excel file
def load_workbook(file_path):
    return xw.Book(file_path)


# Loads sheet in workbook
def select_sheet(workbook, sheet_name):
    return workbook.sheets[sheet_name]


# Saves file
def save(workbook, file_path):
    workbook.save(file_path)
    workbook.close()


# Close app
def close_app(app):
    app.quit()


# Returns the value of a cell in a given worksheet
def get_cell_value(worksheet, row, column):
    # return worksheet.range(f"{to_letter(column)}{row}").value
    return worksheet.range(row, column).value


# Writes into a cell of a given worksheet
def write_to_cell(worksheet, row, column, value):
    # worksheet.range(f"{to_letter(column)}{row}").value = value
    worksheet.range(row, column).value = value


# Adds fill to cell using rgb
def set_fill(worksheet, row: int, column: int, r: int, g: int, b: int):
    # Selected cell
    cell = worksheet.range(row, column)
    # Add fill
    cell.api.Interior.Color = xw.utils.rgb_to_int((r, g, b))


# Converts number to letter
def to_letter(column_int: int):
    start_index = 1

    letter = ''
    while column_int > 25 + start_index:
        letter += chr(65 + int((column_int - start_index) / 26) - 1)
        column_int = column_int - (int((column_int - start_index) / 26)) * 26

    letter += chr(65 - start_index + (int(column_int)))

    return letter
