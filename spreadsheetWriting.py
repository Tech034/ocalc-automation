# ----- Imports ----- #
import openpyxl
import xlwings as xw  # <-- For inserting and deleting rows
import helper.excelFunctions as excel


# Adds pole data to spreadsheet
def add_pole_data(worksheet, pole_dict, insert_from):
    # Insert new cells
    insert_to = insert_from + 20
    insert_row_with_merge(worksheet, insert_from + 1, insert_to)

    # Write in data before make ready
    excel.write_to_cell(worksheet=worksheet, row=insert_from, column=1, value=pole_dict['Pole Number'])  # A17
    excel.write_to_cell(worksheet=worksheet, row=insert_from, column=2, value=pole_dict['Latitude'])  # B17
    excel.write_to_cell(worksheet=worksheet, row=insert_from, column=3, value=pole_dict['Longitude'])  # C17
    excel.write_to_cell(worksheet=worksheet, row=insert_from, column=4, value=pole_dict['Pole Ht/ Class'])  # D17
    excel.write_to_cell(worksheet=worksheet, row=insert_from, column=6, value=pole_dict['Pole Owner'])  # F17
    excel.write_to_cell(worksheet=worksheet, row=insert_from, column=7,
                        value=pole_dict['Proposed Riser (Yes/No) & Qty'])  # G17 <--- Default until for now
    excel.write_to_cell(worksheet=worksheet, row=insert_from, column=8,
                        value=pole_dict['Proposed Guy (Yes/No) & Qty'])  # H17 <--- Default until for now
    excel.write_to_cell(worksheet=worksheet, row=insert_from, column=9, value=pole_dict['Existing Loading'])  # I17
    excel.write_to_cell(worksheet=worksheet, row=insert_from, column=10, value=pole_dict['Proposed Loading'])  # J17
    # K17
    excel.write_to_cell(worksheet=worksheet, row=insert_from, column=11,
                        value=pole_dict['Construction Grade of Analysis'])

    # Adds attacher descriptions and existing heights
    row_num = insert_from
    for attach_description in pole_dict['Make Ready Data']['Attacher Description']:
        # Locate arrays
        existing_arr = pole_dict['Make Ready Data']['Existing Attach Dict'][attach_description]
        proposed_arr = pole_dict['Make Ready Data']['Proposed Attach Dict'][attach_description]

        # If arrays are the exact same
        if check_same_elements(existing_arr, proposed_arr):

            # Adds number to end
            if len(existing_arr) > 1:
                num = 1
                for height in existing_arr:
                    excel.write_to_cell(worksheet=worksheet, row=row_num, column=12,
                                        value=f"{attach_description} {num}")  # Column L
                    excel.write_to_cell(worksheet=worksheet, row=row_num, column=13, value=height)  # Column M
                    num += 1
                    row_num += 1

            # Doesn't add number to end
            else:
                for height in existing_arr:
                    excel.write_to_cell(worksheet=worksheet, row=row_num, column=12,
                                        value=attach_description)  # Column L
                    excel.write_to_cell(worksheet=worksheet, row=row_num, column=13, value=height)  # Column M
                    row_num += 1

        # If arrays are the same length
        elif len(existing_arr) == len(proposed_arr):
            for index in range(len(existing_arr)):
                existing_height = existing_arr[index]
                proposed_height = proposed_arr[index]

                # Add only existing
                if existing_height == proposed_height:
                    if len(existing_arr) > 1:
                        excel.write_to_cell(worksheet=worksheet, row=row_num, column=12,
                                            value=f"{attach_description} {index + 1}")  # Column L
                        excel.write_to_cell(worksheet=worksheet, row=row_num, column=13,
                                            value=existing_height)  # Column M
                    else:
                        excel.write_to_cell(worksheet=worksheet, row=row_num, column=12,
                                            value=attach_description)  # Column L
                        excel.write_to_cell(worksheet=worksheet, row=row_num, column=13,
                                            value=existing_height)  # Column M

                # Add existing and proposed
                else:
                    if len(existing_arr) > 1:
                        excel.write_to_cell(worksheet=worksheet, row=row_num, column=12,
                                            value=f"{attach_description} {index + 1}")  # Column L
                        excel.write_to_cell(worksheet=worksheet, row=row_num, column=13,
                                            value=existing_height)  # Column M
                        excel.write_to_cell(worksheet=worksheet, row=row_num, column=14,
                                            value=proposed_height)  # Column N

                        # Adds red fill if power is moving and pole isn't
                        if 'crossarm' in attach_description.lower() and 'nwe' in attach_description.lower() and not \
                        pole_dict['New Pole'] and not pole_dict['Pole Moved']:
                            excel.set_fill(worksheet=worksheet, row=row_num, column=14, r=255, g=0, b=0)
                    else:
                        excel.write_to_cell(worksheet=worksheet, row=row_num, column=12,
                                            value=attach_description)  # Column L
                        excel.write_to_cell(worksheet=worksheet, row=row_num, column=13,
                                            value=existing_height)  # Column M
                        excel.write_to_cell(worksheet=worksheet, row=row_num, column=14,
                                            value=proposed_height)  # Column N

                        # Adds red fill if power is moving and pole isn't
                        if 'crossarm' in attach_description.lower() and 'nwe' in attach_description.lower() and not \
                                pole_dict['New Pole'] and not pole_dict['Pole Moved']:
                            excel.set_fill(worksheet=worksheet, row=row_num, column=14, r=255, g=0, b=0)

                # Increment row
                row_num += 1

        # If arrays are a different length
        else:
            # Add all existing heights
            num = 1
            for height in existing_arr:
                if len(existing_arr) > 1:
                    excel.write_to_cell(worksheet=worksheet, row=row_num, column=12,
                                        value=f"{attach_description} {num}")  # Column L
                    excel.write_to_cell(worksheet=worksheet, row=row_num, column=13, value=height)  # Column M
                else:
                    excel.write_to_cell(worksheet=worksheet, row=row_num, column=12,
                                        value=attach_description)  # Column L
                    excel.write_to_cell(worksheet=worksheet, row=row_num, column=13,
                                        value=height)  # Column M

                num += 1
                row_num += 1

            # Adds all proposed heights to the same cell
            if not check_same_elements(existing_arr, proposed_arr):
                proposed = ""
                for height in pole_dict['Make Ready Data']['Proposed Attach Dict'][attach_description]:
                    proposed += f"{height}  "
                excel.write_to_cell(worksheet=worksheet, row=row_num, column=12, value=attach_description)  # Column L
                excel.write_to_cell(worksheet=worksheet, row=row_num, column=14, value=proposed)  # Column N

                # Adds red fill if power is moving and pole isn't
                if 'crossarm' in attach_description.lower() and 'nwe' in attach_description.lower() and not pole_dict[
                        'New Pole'] and not pole_dict['Pole Moved']:
                    excel.set_fill(worksheet=worksheet, row=row_num, column=14, r=255, g=0, b=0)

                row_num += 1

        # # Adds attacher description for each existing attachment
        # for height in existing_arr:
        #     excel.write_to_cell(worksheet=worksheet, row=row_num, column=12, value=attach_description)  # Column L
        #     excel.write_to_cell(worksheet=worksheet, row=row_num, column=13, value=height)  # Column M
        #     row_num += 1
        #
        # # Adds all proposed heights to the same cell
        # if not check_same_elements(existing_arr, proposed_arr):
        #     proposed = ""
        #     for height in pole_dict['Make Ready Data']['Proposed Attach Dict'][attach_description]:
        #         proposed += f"{height}  "
        #     excel.write_to_cell(worksheet=worksheet, row=row_num, column=12, value=attach_description)  # Column L
        #     excel.write_to_cell(worksheet=worksheet, row=row_num, column=14, value=proposed)  # Column N
        #     row_num += 1

    # Deletes extra rows for pole
    delete_from = find_next_open_row(worksheet, 12)
    delete_to = insert_from + 20 + 9 - 1
    delete_row_with_merge(worksheet, delete_from, delete_to)


# Returns the row number with empty cell for specified column
def find_next_open_row(worksheet, column_index):  # <-- MOST LIKELY USE '12'
    next_open_row = worksheet.cells.last_cell.row + 1

    for row in range(17, worksheet.cells.last_cell.row + 1):
        cell_value = excel.get_cell_value(worksheet=worksheet, row=row, column=column_index)
        if cell_value is None:
            next_open_row = row
            break

    return next_open_row


# Adds rows to spreadsheet
def insert_row_with_merge(worksheet, from_index, to_index):
    # Select range
    selected_range = f"{from_index}:{to_index}"

    # Insert selected range
    worksheet.range(selected_range).api.Insert()


# Removes rows from spreadsheet
def delete_row_with_merge(worksheet, from_index, to_index):
    # Select range
    selected_range = f"{from_index}:{to_index}"

    # Insert selected range
    worksheet.range(selected_range).api.Delete()


# Checks if two arrays contain the same elements
def check_same_elements(arr1, arr2):
    return set(arr1) == set(arr2)
