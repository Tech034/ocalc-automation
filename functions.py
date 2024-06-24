# Returns inches measurement to feet (whole number)
def inches_to_feet(inches):
    return round(float(inches) / 12)


# Returns attachment height minus bury depth in inches to feet and inches
def to_feet_and_inches(attach_path, depth, item_key):
    height = float(get_item(attach_path, item_key)) - float(depth)
    feet = int(height / 12)
    inches = int(round(height % 12, 0))
    if inches == 12:
        feet += 1
        inches = 0
    string = f"{feet}' "
    string += f'{inches}"'
    return string


# Returns item for pole attachment
def get_item(attach_path, item_key):
    for x in attach_path[0]:  # [0] goes into attributes
        if x.attrib['NAME'] == item_key:
            return x.text


# Returns true if the heights of the attachments are the exact same (Ex: if both have a height of 516.572275445563)
def compare_attach(attach_path1, attach_path2, item_key):
    if get_item(attach_path1, item_key) == get_item(attach_path2, item_key):
        return True
    else:
        return False


# returns true if given value is contained in attach list path for a certain key
def contains(attach_list_path, value, item_key):
    for index, attach_path in enumerate(attach_list_path, start=1):  # Skips the LoadCase
        if get_item(attach_path, item_key) == value:
            return True
    return False
