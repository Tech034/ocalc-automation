# ----- Imports ----- #
import xml.etree.ElementTree as ET
import functions as f
import re


def get_pole_data(existing_file, proposed_file):
    # ----- Open XML ----- #
    existing_tree = ET.parse(existing_file)
    existing_root = existing_tree.getroot()
    proposed_tree = ET.parse(proposed_file)
    proposed_root = proposed_tree.getroot()

    pole_dict = {
        "Pole Number": extract_number(existing_file),
        "Latitude": None,  # float(existing_root[0][0][1].text)
        "Longitude": None,  # float(existing_root[0][0][2].text)
        "Pole Ht/ Class": None,  # f"{f.inches_to_feet(existing_root[0][1][0][0][35].text)}/{existing_root[0][1][0][0][6].text}"
        "Proposed Ht/ Class": None,  # f"{f.inches_to_feet(proposed_root[0][1][0][0][35].text)}/{proposed_root[0][1][0][0][6].text}"
        "Span Length": None,
        "Pole Owner": "NWE",
        "Proposed Riser (Yes/No) & Qty": None,
        "Proposed Guy (Yes/No) & Qty": None,
        "Existing Loading": None,  # f"{existing_root[0][2][0][0][0].text}%
        "Proposed Loading": None,  # f"{proposed_root[0][2][0][0][0].text}%"
        "Construction Grade of Analysis": "C",  # existing_root[0][1][0][1][0][0][18].text
        "Make Ready Data": {
            "Attacher Description": [],
            "Existing Attach Dict": {},
            "Proposed Attach Dict": {},
        },
        "Communication Make Ready": None,
        "Utility Make Ready": None,
        "Odd Attachments": [],
        "New Pole": new_pole(existing_root, proposed_root),
        "Pole Moved": pole_moved(existing_root, proposed_root),
    }

    # Fill out dict
    try:
        pole_dict["Latitude"] = float(existing_root[0][0][1].text)
    except IndexError:
        pass

    try:
        pole_dict["Longitude"] = float(existing_root[0][0][2].text)
    except IndexError:
        pass

    try:
        pole_dict["Pole Ht/ Class"] = f"{f.inches_to_feet(existing_root[0][1][0][0][35].text)}/{existing_root[0][1][0][0][6].text}"
    except IndexError:
        pass

    try:
        pole_dict["Proposed Ht/ Class"] = f"{f.inches_to_feet(proposed_root[0][1][0][0][35].text)}/{proposed_root[0][1][0][0][6].text}"
    except IndexError:
        pass

    try:
        pole_dict["Existing Loading"] = f"{existing_root[0][2][0][0][0].text}%"
    except IndexError:
        pass

    try:
        pole_dict["Proposed Loading"] = f"{proposed_root[0][2][0][0][0].text}%"
    except IndexError:
        pass

    # Fills out existing attach dict and proposed attach dict inside pole dict
    attach_dict(pole_dict, existing_file, True)
    attach_dict(pole_dict, proposed_file, False)

    # Set proposed riser
    existing_riser_count = int(pole_dict['Make Ready Data']['Existing Attach Dict']['Riser Count'])
    proposed_riser_count = int(pole_dict['Make Ready Data']['Proposed Attach Dict']['Riser Count'])
    num_proposed_risers = proposed_riser_count - existing_riser_count
    if num_proposed_risers > 0:
        pole_dict['Proposed Riser (Yes/No) & Qty'] = f"Yes/{num_proposed_risers}"
    else:
        pole_dict['Proposed Riser (Yes/No) & Qty'] = "No"

    # Set proposed guy
    existing_guy_count = int(pole_dict['Make Ready Data']['Existing Attach Dict']['Guy Count'])
    proposed_guy_count = int(pole_dict['Make Ready Data']['Proposed Attach Dict']['Guy Count'])
    num_proposed_guys = proposed_guy_count - existing_guy_count
    if num_proposed_guys > 0:
        pole_dict['Proposed Guy (Yes/No) & Qty'] = f"Yes/{num_proposed_guys}"
    else:
        pole_dict['Proposed Guy (Yes/No) & Qty'] = "No"

    # Sort data
    if 'Drip Loop' in pole_dict['Make Ready Data']['Attacher Description']:
        swap_last_strings(pole_dict['Make Ready Data']['Attacher Description'], 'Drip Loop', 'NWE Secondary Drop')

    return pole_dict


# Fills out given 'Existing Attach Dict' or 'Proposed Attach Dict' using given file
def attach_dict(pole_dict, file_name, is_existing: bool):
    # ----- Open XML ----- #
    tree = ET.parse(file_name)
    root = tree.getroot()

    # How deep pole is buried in inches
    bury_depth = root[0][1][0][0][11].text

    # Riser and guy hook count
    riser_count = 0
    guy_count = 0

    # Set to fill existing or proposed dictionary
    if is_existing:
        attach_dict_key = 'Existing Attach Dict'
    else:
        attach_dict_key = 'Proposed Attach Dict'

    # # ------------------------------------------------ Temp ------------------------------------------------- #
    # # Adds bury depth as an attachment so that it appears in the spreadsheet
    # b_height = float(bury_depth)
    # b_feet = int(b_height / 12)
    # b_inches = int(round(b_height % 12, 0))
    # if b_inches == 12:
    #     b_feet += 1
    #     b_inches = 0
    # b_string = f"{b_feet}' "
    # b_string += f'{b_inches}"'
    # add_attachment(pole_dict, 'Bury Depth', b_string, attach_dict_key)
    # # ------------------------------------------------ Temp ------------------------------------------------- #

    # Loops through existing attachments
    for index, attachment in enumerate(root[0][1][0][1][0:], start=0):  # Skips the LoadCase

        # Stores the name of the attachment
        attachment_name = ""
        height = None

        # Checks for tag
        if attachment.tag == 'Crossarm':

            description = ""

            try:
                for insulator in attachment[1]:
                    if insulator[1][0].tag == 'Span':
                        if 'primary' in f.get_item(insulator[1][0], 'SpanType').lower():
                            description = 'primary'
                            break
                        elif 'secondary' in f.get_item(insulator[1][0], 'SpanType').lower():
                            description = 'secondary'
                            break
                        else:
                            description = f.get_item(attachment, 'Description').lower()
                    else:
                        description = f.get_item(attachment, 'Description').lower()
            except (AttributeError, IndexError):
                description = f.get_item(attachment, 'Description').lower()

            # If the owner is NWE
            if f.get_item(attachment, 'Owner') == 'NWE':

                if 'primary' in description or 'pupi tb2000120' in description:
                    attachment_name = 'NWE Primary Crossarm'
                elif 'secondary' in description:
                    attachment_name = 'NWE Secondary Crossarm'

                    # Add Crossarm Brace
                    attachment_name2 = 'Crossarm Brace'
                    height2 = None
                    if attachment_name2 not in pole_dict['Make Ready Data']['Attacher Description']:
                        add_attachment_with_check(pole_dict, attachment_name2, height2, attach_dict_key, attachment)
                    # Has to check separately because len cant be checked unless first 'if'  has already happened
                    elif len(pole_dict['Make Ready Data'][attach_dict_key][attachment_name2]) == 0:
                        add_attachment_with_check(pole_dict, attachment_name2, height2, attach_dict_key, attachment)

                elif 'secondary spool rack' in description:
                    attachment_name = 'NWE Secondary Spool Rack'
                elif 'vertical standoff bracket' in description:  # may not be needed
                    attachment_name = f"{f.get_item(attachment, 'Owner')} Vertical Standoff Bracket"
                elif 'crossarm' in description:  # may not be needed
                    attachment_name = f"{f.get_item(attachment, 'Owner')} Crossarm"
                else:
                    attachment_name = f.get_item(attachment, 'Description')
                    # Adds to list of odd attachments
                    if attachment_name not in pole_dict["Odd Attachments"]:
                        pole_dict["Odd Attachments"].append(attachment_name)
                    # raise ValueError("Error: Crossarm did not contain 'primary', 'secondary', "
                    #                  "'secondary spool rack', or 'vertical standoff bracket'")

            else:  # If the owner is not NWE

                if 'vertical standoff bracket' in description:
                    attachment_name = f"{f.get_item(attachment, 'Owner')} Vertical Standoff Bracket"
                elif 'crossarm' in description:
                    attachment_name = f"{f.get_item(attachment, 'Owner')} Crossarm"
                else:
                    attachment_name = f"{f.get_item(attachment, 'Owner')} {f.get_item(attachment, 'Description')}"
                    # Adds to list of odd attachments
                    if attachment_name not in pole_dict["Odd Attachments"]:
                        pole_dict["Odd Attachments"].append(attachment_name)
                    # raise ValueError("Error: Crossarm did not contain "
                    #                  "'secondary spool rack', or 'vertical standoff bracket'")

            # Update height
            height = f.to_feet_and_inches(attachment, bury_depth, item_key='CoordinateZ')

            # Add attachment
            add_attachment_with_check(pole_dict, attachment_name, height, attach_dict_key, attachment)

        elif attachment.tag == 'Insulator':

            # Checks for type of insulator attachment
            if f.get_item(attachment, 'Type') == 'Spool':

                description = ""
                contains_service = False

                # Creates name for insulator spool attachment
                try:
                    for span in attachment[1]:
                        description = f.get_item(span, 'SpanType').lower()
                        if 'neutral' in description or 'secondary' in description:
                            break
                        elif 'service' in description:
                            contains_service = True
                        else:
                            description = f.get_item(attachment, 'Description').lower()
                except IndexError:
                    description = f.get_item(attachment, 'Description').lower()

                if contains_service and 'neutral' not in description and 'secondary' not in description:
                    description = 'service'

                # try:
                #     description = f.get_item(attachment[1][0], 'SpanType').lower()
                #     if 'primary' in description or 'secondary' in description:
                #         pass
                #     else:
                #         description = f.get_item(attachment, 'Description').lower()
                # except IndexError:
                #     description = f.get_item(attachment, 'Description').lower()

                if 'neutral' in description:
                    attachment_name = 'NWE Neutral'
                elif 'secondary' in description or 'service' in description:
                    attachment_name = 'NWE Secondary Drop'

                    # Add Drip Loop
                    attachment_name2 = 'Drip Loop'
                    height2 = None
                    if attachment_name2 not in pole_dict['Make Ready Data']['Attacher Description']:
                        add_attachment_with_check(pole_dict, attachment_name2, height2, attach_dict_key, attachment)
                    # Has to check separately because len cant be checked unless first 'if'  has already happened
                    elif len(pole_dict['Make Ready Data'][attach_dict_key][attachment_name2]) == 0:
                        add_attachment_with_check(pole_dict, attachment_name2, height2, attach_dict_key, attachment)

                else:
                    attachment_name = f.get_item(attachment, 'Description')
                    # Adds to list of odd attachments
                    if attachment_name not in pole_dict["Odd Attachments"]:
                        pole_dict["Odd Attachments"].append(attachment_name)
                    # raise ValueError("Error: Insulator Spool did not contain 'neutral' or 'secondary'")

            elif f.get_item(attachment, 'Type') == 'Bolt' or f.get_item(attachment, 'Type') == 'J-Hook':

                # Creates name for insulator bolt attachment
                if f.get_item(attachment, 'Owner') == 'NWE':
                    attachment_name = 'NWE Fiber'
                elif f.get_item(attachment, 'Owner') == 'TDS':
                    attachment_name = 'TDS'
                elif f.get_item(attachment, 'Owner') == 'Spectrum':
                    attachment_name = 'Spectrum CATV'
                elif f.get_item(attachment, 'Owner') == 'Ziply':
                    attachment_name = 'Ziply Fiber'
                elif f.get_item(attachment, 'Owner') == 'Lumen':
                    attachment_name = 'Lumen'
                else:
                    attachment_name = f"{f.get_item(attachment, 'Owner')} {f.get_item(attachment, 'Type')}"
                    # Adds to list of odd attachments
                    if attachment_name not in pole_dict["Odd Attachments"]:
                        pole_dict["Odd Attachments"].append(attachment_name)
                    # raise ValueError(
                    #     "Error: Insulator Bolt did not contain 'NWE', 'TDS', 'Spectrum', 'Ziply', or 'Lumen'")

            elif f.get_item(attachment, 'Type') == 'Pin':

                # Creates name for pin attachment
                attachment_name = f"{f.get_item(attachment, 'Owner')} Pin"

            elif f.get_item(attachment, 'Type') == 'Deadend':

                # Creates name for pin attachment
                attachment_name = f"{f.get_item(attachment, 'Owner')} Deadend"

            else:
                attachment_name = f"{f.get_item(attachment, 'Owner')} {f.get_item(attachment, 'Type')}"
                # Adds to list of odd attachments
                if attachment_name not in pole_dict["Odd Attachments"]:
                    pole_dict["Odd Attachments"].append(attachment_name)
                # raise ValueError("Error: Insulator was not of type 'Spool' or 'Bolt'")

            # Update height
            height = f.to_feet_and_inches(attachment, bury_depth, item_key='CoordinateZ')

            # Add attachment
            add_attachment_with_check(pole_dict, attachment_name, height, attach_dict_key, attachment)

        elif attachment.tag == 'Streetlight':

            # Creates name for streetlight attachment
            if 'streetlight' in f.get_item(attachment, 'Description').lower():
                attachment_name = 'Street Light Drip'

                # Add Streetlight
                attachment_name2 = 'Street Light'
                height2 = f.to_feet_and_inches(attachment, bury_depth, item_key='CoordinateZ')
                add_attachment_with_check(pole_dict, attachment_name2, height2, attach_dict_key, attachment)

            else:
                attachment_name = f.get_item(attachment, 'Description')
                # Adds to list of odd attachments
                if attachment_name not in pole_dict["Odd Attachments"]:
                    pole_dict["Odd Attachments"].append(attachment_name)

                # Update height
                height = f.to_feet_and_inches(attachment, bury_depth, item_key='CoordinateZ')
                # raise ValueError("Error: Streetlight did not contain 'streetlight' text")

            # Add attachment
            add_attachment_with_check(pole_dict, attachment_name, height, attach_dict_key, attachment)

        elif attachment.tag == 'Riser':

            # Creates name for riser attachment
            if 'riser' in f.get_item(attachment, 'Description').lower():
                attachment_name = f"{f.get_item(attachment, 'Owner')} Riser"
                riser_count += 1
            else:
                attachment_name = f.get_item(attachment, 'Description')
                # Adds to list of odd attachments
                if attachment_name not in pole_dict["Odd Attachments"]:
                    pole_dict["Odd Attachments"].append(attachment_name)
                # raise ValueError("Error: Riser did not contain 'riser' text")

            # Update height
            height = f.to_feet_and_inches(attachment, depth=0, item_key='LengthAboveGLInInches')

            # Add attachment
            add_attachment_with_check(pole_dict, attachment_name, height, attach_dict_key, attachment)

        elif attachment.tag == 'PowerEquipment':

            # Creates name for PowerEquipment attachment
            if f.get_item(attachment, 'Type') == 'Transformer':
                attachment_name = 'NWE Base Transformer'

                # Add Top Transformer
                attachment_name2 = 'NWE Top Transformer'
                height2 = f.to_feet_and_inches(attachment, bury_depth, item_key='CoordinateZ')
                unit_height = int(f.get_item(attachment, 'HeightInInches'))
                # Get top of Transformer
                height2 = add_inches(height2, unit_height / 2)
                add_attachment_with_check(pole_dict, attachment_name2, height2, attach_dict_key, attachment)

                # Add Base Transformer
                try:
                    height = add_inches(height2, unit_height * -1)  # Gets base by subtracting transformer height
                    base_array = pole_dict['Make Ready Data'][attach_dict_key]['NWE Base Transformer']
                    top_array = pole_dict['Make Ready Data'][attach_dict_key]['NWE Top Transformer']
                    if len(base_array) < len(top_array):
                        add_attachment_with_check(pole_dict, attachment_name, height, attach_dict_key, attachment)
                except KeyError:
                    add_attachment_with_check(pole_dict, attachment_name, height, attach_dict_key, attachment)

            else:
                attachment_name = f.get_item(attachment, 'Description')
                # Adds to list of odd attachments
                if attachment_name not in pole_dict["Odd Attachments"]:
                    pole_dict["Odd Attachments"].append(attachment_name)

                # Update height and add
                height = f.to_feet_and_inches(attachment, bury_depth, item_key='CoordinateZ')
                add_attachment_with_check(pole_dict, attachment_name, height, attach_dict_key, attachment)

                # raise ValueError("Error: PowerEquipment did not contain 'transformer' text")

        elif attachment.tag == 'Anchor':

            # Adds to guy count
            guy_count += 1

            # Checks for attachment type
            try:
                if f.get_item(attachment[1][0], 'Type') == 'Sidewalk':  # Note: attachment[1][0] has 'Type' + 'CoordinateZ'
                    attachment_name = f"{f.get_item(attachment, 'Owner')} Sidewalk Guy"
                elif f.get_item(attachment[1][0], 'Type') == 'Down':
                    attachment_name = f"{f.get_item(attachment, 'Owner')} Downguy"
                elif f.get_item(attachment[1][0], 'Type') == 'Span/Head':
                    attachment_name = f"{f.get_item(attachment, 'Owner')} Span/Head"
                    # Cancels guy count
                    guy_count -= 1
                else:
                    attachment_name = f"{f.get_item(attachment, 'Owner')} {f.get_item(attachment[1][0], 'Type')}"
                    # Adds to list of odd attachments
                    if attachment_name not in pole_dict["Odd Attachments"]:
                        pole_dict["Odd Attachments"].append(attachment_name)
                    # Cancels guy count
                    guy_count -= 1

                # Update height
                height = f.to_feet_and_inches(attachment[1][0], bury_depth, item_key='CoordinateZ')

            except IndexError:
                attachment_name = f"Anchor"

                # Update height
                height = 'N/A'

            # Add attachment
            add_attachment_with_check(pole_dict, attachment_name, height, attach_dict_key, attachment)

        elif attachment.tag == 'LoadCase':

            # Makes sure LoadCase is ignored
            pass

        else:
            attachment_name = f"{attachment.tag} Attachment"

            # Add attachment
            add_attachment_with_check(pole_dict, attachment_name, height, attach_dict_key, attachment)

            # Adds to list of odd attachments
            if attachment_name not in pole_dict["Odd Attachments"]:
                pole_dict["Odd Attachments"].append(attachment_name)
            # raise ValueError("Tag did not match any of the following: Crossarm, Insulator, Streetlight, Riser,"
            #                  " PowerEquipment, Anchor, LoadCase")

    # Remove Deadend attachments within 4" of crossarms
    if 'NWE Deadend' in pole_dict['Make Ready Data']['Attacher Description']:
        remove_array = []
        current_array = pole_dict['Make Ready Data'][attach_dict_key]['NWE Deadend']
        for index, height in enumerate(current_array):
            for attachment_dict in pole_dict['Make Ready Data']['Attacher Description']:
                if 'crossarm' in attachment_dict.lower():
                    for crossarm_height in pole_dict['Make Ready Data'][attach_dict_key][attachment_dict]:
                        if within_four_inches(height, crossarm_height):
                            remove_array.append(height)

        # Remove items
        pole_dict['Make Ready Data'][attach_dict_key]['NWE Deadend'] = [i for i in current_array if i not in remove_array]

    # Adds the number of risers and guys to dictionary
    pole_dict['Make Ready Data'][attach_dict_key]['Riser Count'] = riser_count
    pole_dict['Make Ready Data'][attach_dict_key]['Guy Count'] = guy_count


# Adds attachment to pole_dict
def add_attachment(pole_dict, attachment_name, height, attach_dict_key):
    # Add attachment to 'Attacher Description' and 'Existing Attach Dict' if not there
    if attachment_name not in pole_dict['Make Ready Data']['Attacher Description'] and attachment_name != '':
        pole_dict['Make Ready Data']['Attacher Description'].append(attachment_name)
        pole_dict['Make Ready Data']['Existing Attach Dict'][attachment_name] = []
        pole_dict['Make Ready Data']['Proposed Attach Dict'][attachment_name] = []

    # Add height to 'Existing_Attach_Dict'
    if attachment_name in pole_dict['Make Ready Data']['Attacher Description']:
        pole_dict['Make Ready Data'][attach_dict_key][attachment_name].append(height)


def add_attachment_with_check(pole_dict, attachment_name, height, attach_dict_key, attachment):
    # Doesn't add attachment within 4" of another attachment unless it is a riser
    if attachment.tag == 'Riser' or height is None:
        add_attachment(pole_dict, attachment_name, height, attach_dict_key)
    else:
        try:
            is_close = False
            for list_height in pole_dict['Make Ready Data'][attach_dict_key][attachment_name]:
                if within_four_inches(list_height, height):
                    is_close = True

            if not is_close:
                add_attachment(pole_dict, attachment_name, height, attach_dict_key)

        except KeyError:
            add_attachment(pole_dict, attachment_name, height, attach_dict_key)


# Adds inches to feet an inches string
def add_inches(feet_inches_str: str, inches: int):
    # Split the feet and inches from the input string
    feet_inches = feet_inches_str.split("'")
    feet = int(feet_inches[0].strip())
    inches_from_feet = int(feet_inches[1].strip().replace('"', ''))

    # Convert feet to inches and add to the inches from the input string
    total_inches = (feet * 12) + inches_from_feet

    # Add the additional inches
    total_inches += inches

    # Convert back to feet and inches
    total_feet = int(total_inches // 12)
    remaining_inches = int(total_inches % 12)

    return f"{total_feet}' {remaining_inches}\""


def within_four_inches(height1, height2):
    try:
        # Split the strings into feet and inches
        feet1, inches1 = int(height1.split("'")[0]), int(height1.split("'")[1].split('"')[0])
        feet2, inches2 = int(height2.split("'")[0]), int(height2.split("'")[1].split('"')[0])

        # Convert everything to inches
        total_inches1 = feet1 * 12 + inches1
        total_inches2 = feet2 * 12 + inches2

        # Check if the difference is within 4 inches
        return abs(total_inches1 - total_inches2) <= 4
    except (ValueError, AttributeError):
        return False


# If bury depth is different on proposed it returns true
def pole_moved(existing_root, proposed_root):
    try:
        if existing_root[0][1][0][0][11].text == proposed_root[0][1][0][0][11].text:
            return False
        else:
            return True
    except IndexError:
        return False


# Returns true if pole is new
def new_pole(existing_root, proposed_root):
    try:
        existing_class = f"{f.inches_to_feet(existing_root[0][1][0][0][35].text)}/{existing_root[0][1][0][0][6].text}"
        proposed_class = f"{f.inches_to_feet(proposed_root[0][1][0][0][35].text)}/{proposed_root[0][1][0][0][6].text}"
        return existing_class != proposed_class
    except IndexError:
        return False


def swap_last_strings(arr, str1, str2):
    """
    This function takes an array and two strings as inputs. It swaps the last occurrences of the two strings in the
    array.

    Parameters:
    arr (list): The list in which the strings are to be swapped.
    str1 (str): The first string to be swapped.
    str2 (str): The second string to be swapped.

    Returns:
    list: The list with the last occurrences of the strings swapped.
    """
    # Find the last occurrences of the strings
    last_str1 = len(arr) - 1 - arr[::-1].index(str1)
    last_str2 = len(arr) - 1 - arr[::-1].index(str2)

    # Swap the last occurrences
    arr[last_str1], arr[last_str2] = arr[last_str2], arr[last_str1]

    return arr


def extract_number(s):
    match = re.search('.*?(\d.*?(?=-Existing|-Proposed))', s)
    if match:
        return match.group(1).lstrip('0') or '0'
    else:
        return None
