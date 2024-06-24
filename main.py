# ----- Imports ----- #
import os
import re
import dataExtraction as extract
import spreadsheetWriting as sheet
import helper.excelFunctions as excel
import highlight as mark
from sort import sort_rows


# ----- Functions ----- #
# Sorts the files regex
def sort_key(filename):
    match = re.search(r'(\d+)-(\d+)(?=-Existing|-Proposed)', filename)
    if match:
        # Return a tuple of integers for sorting
        return int(match.group(1)), int(match.group(2))
    else:
        return float('inf'), float('inf')


# Gets all the existing and proposed files together
def get_pairs(directory):
    # Create a dictionary to hold the pairs_dict
    pairs_dict = {}

    for filename in sorted(os.listdir(directory), key=sort_key):
        # Check if the file has the correct extension
        if filename.endswith('.pplx'):
            # Extract the base name of the file
            base_name = filename.replace('-Proposed', '').replace('-Existing', '').replace('Pole_', '').replace('_pplx', '').replace('_', '')

            # If the base name is already in the dictionary, add the file to the pair
            if base_name in pairs_dict:
                pairs_dict[base_name].append(os.path.join(directory, filename))
            # If the base name is not in the dictionary, create a new pair
            else:
                pairs_dict[base_name] = [os.path.join(directory, filename)]

    return pairs_dict


# ----- Code ----- #
# Using 'template.xlsx' to create 'output.xlsx' in output folder
app = excel.set_app(is_visible=False)
wb = excel.load_workbook(file_path='template.xlsx')
ws = excel.select_sheet(workbook=wb, sheet_name='Sheet1')

# Organize files into a dictionary
directory_path = 'xml'  # Replace with your directory
pairs = get_pairs(directory_path)

# Create pole dict
odd_attachments = []

# Now you can iterate over the pairs
for base, pair in pairs.items():
    if len(pair) == 2:
        # Separate file paths
        existing_file_path, proposed_file_path = pair

        # Create a dictionary with the pole data
        pole_dict = extract.get_pole_data(existing_file_path, proposed_file_path)

        # Add to odd attachments
        odd_attachments.extend(i for i in pole_dict["Odd Attachments"] if i not in odd_attachments)

        # Open sheet to get next open row
        open_row = sheet.find_next_open_row(ws, 12)

        # Use 'pole_dict' and the next open row in spreadsheet to insert data
        sheet.add_pole_data(ws, pole_dict, open_row)

        # Update poles in console
        print(f"{extract.extract_number(proposed_file_path)} was added to spreadsheet")
    else:
        print(f'Warning: {base} does not have a pair.')
        # # --------- For existing only files --------- #
        # # Create existing
        # existing_file_path = 'xml/' + base.replace('.pplx', '') + '-Existing.pplx'
        # # existing_file_path = 'xml/' + 'Pole_' + base.replace('.pplx', '') + '_pplx.pplx'
        #
        # # Create a dictionary with the pole data
        # pole_dict = extract.get_pole_data(existing_file_path, existing_file_path)
        #
        # # Add to odd attachments
        # odd_attachments.extend(i for i in pole_dict["Odd Attachments"] if i not in odd_attachments)
        #
        # # Open sheet to get next open row
        # open_row = sheet.find_next_open_row(ws, 12)
        #
        # # Use 'pole_dict' and the next open row in spreadsheet to insert data
        # sheet.add_pole_data(ws, pole_dict, open_row)
        #
        # # Update poles in console
        # print(f"{extract.extract_number(existing_file_path)} was added to spreadsheet")
        # # ------------------------------------------- #

# Sort attachments
sort_rows(ws)

# Highlight cells
mark.highlight(ws, odd_attachments)

# Save and close workbook in new location
excel.save(workbook=wb, file_path='output/output.xlsx')
excel.close_app(app=app)
