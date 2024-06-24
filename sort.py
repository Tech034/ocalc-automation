# ----- Imports ----- #
from typing import Dict, List
import xlwings as xw
import re
import helper.excelFunctions as excel


def move_attachment(data_list: List[Dict], attachment_to_move: str, attachment_to_move_after: str):
    """Moves a given attachment after another given attachment in list"""
    saved_attachment = None

    # Loop through the attachments list in reverse order
    for attachment in reversed(data_list):
        if attachment['name'] == attachment_to_move:
            # Save the attachment dictionary
            saved_attachment = attachment

            # Remove the attachment from the list
            data_list.remove(attachment)

    # Loop through the attachments list again in reverse order
    for i in range(len(data_list) - 1, -1, -1):
        attachment = data_list[i]
        if attachment_to_move_after in attachment['name']:
            # Insert the saved attachment dictionary after this attachment
            data_list.insert(i + 1, saved_attachment)
            break

    # Return the updated attachments list
    return data_list


def sort_rows(sheet):
    # Initialize first section
    print("Sorting by height...")
    start_row = 17
    is_attachment_name = True

    # Loop through sections
    while is_attachment_name:

        # Initialize section data
        row = start_row
        section = sheet.range(f'A{row}').value
        data = []

        # Loop through section rows
        while (sheet.range(f'A{row}').value == section or sheet.range(f'A{row}').value is None) and sheet.range(f'L{row}').value is not None:

            # Store moving attachment data
            attachment = {
                'name': sheet.range(f'L{row}').value,
                'existing_height': sheet.range(f'M{row}').value,
                'proposed_height': sheet.range(f'N{row}').value
            }

            # Get height to convert to inches
            height = "0' 0\""
            if attachment['existing_height'] is not None:
                height = attachment['existing_height']
            elif attachment['proposed_height'] is not None:
                height = attachment['proposed_height']

            # Try to match the feet and inches in the value
            match = re.match(r"(\d+)' (\d+)\"", height)

            # If the match was successful, convert the feet and inches to total inches
            if match:
                feet, inches = match.groups()
                total_inches = int(feet) * 12 + int(inches)
            else:
                # If the match was not successful, set the total inches to a very low value
                total_inches = -1

            # Add inches to dictionary
            attachment['inches'] = total_inches

            # Add attachment to data
            data.append(attachment)

            # Increment
            row += 1

        # Sort attachments in data from highest to lowest
        data.sort(key=lambda x: x['inches'], reverse=True)

        # Move 'Drip Loop' behind last 'Secondary Drop'
        data = move_attachment(data, 'Drip Loop', 'Secondary Drop')

        # Move 'Street Light Drip' behind last 'Street Light'
        data = move_attachment(data, 'Street Light Drip', 'Street Light')

        # Move 'Street Light Drip' behind last 'Street Light'
        data = move_attachment(data, 'Crossarm Brace', 'NWE Secondary Crossarm')

        # Write attachment into the section in the new order
        for index, attachment in enumerate(data):
            sheet.range(f'L{start_row + index}').value = attachment['name']
            sheet.range(f'M{start_row + index}').value = attachment['existing_height']
            sheet.range(f'N{start_row + index}').value = attachment['proposed_height']

        # Set the next sections start row
        start_row = row

        # End sections loop
        if sheet.range(f'L{row}').value is None:
            is_attachment_name = False

    # Show that it's done sorting
    print("Sorting Complete")


# # ----- Test ----- #
# # Using 'template.xlsx' to create 'output.xlsx' in output folder
# app = excel.set_app(is_visible=False)
# wb = excel.load_workbook(file_path='output/output.xlsx')
# ws = excel.select_sheet(workbook=wb, sheet_name='Sheet1')
#
# # Method being tested
# sort_rows(ws)
#
# # Save and close workbook in new location
# excel.save(workbook=wb, file_path='output/output_2.xlsx')
# excel.close_app(app=app)
