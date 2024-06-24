def add_inches(feet_inches_str, inches):
    # Split the feet and inches from the input string
    feet_inches = feet_inches_str.split("'")
    feet = int(feet_inches[0].strip())
    inches_from_feet = int(feet_inches[1].strip().replace('"', ''))

    # Convert feet to inches and add to the inches from the input string
    total_inches = (feet * 12) + inches_from_feet

    # Add the additional inches
    total_inches += inches

    # Convert back to feet and inches
    total_feet = total_inches // 12
    remaining_inches = total_inches % 12

    return f"{total_feet}' {remaining_inches}\""

# Example usage:
print(add_inches("34' 6\"", 16))  # Outputs: "35' 10\""

