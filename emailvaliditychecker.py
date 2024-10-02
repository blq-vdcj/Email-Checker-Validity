import openpyxl
import re
import os

# Load the Excel file
file = "email.xlsx"  # Name of the file
try:
    wb = openpyxl.load_workbook(file)
    sheet = wb.active
except Exception as e:
    with open("error_log.txt", "w") as error_file:
        error_file.write(f"An error occurred while loading the file: {e}\n")
    print(f"An error occurred while loading the file: {e}. Details saved in error_log.txt.")
    input("Press Enter to close the window.")
    exit()

# Function to check the validity of email addresses
def is_email_valid(email):
    pattern = r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
    
    # Check if the email matches the pattern, does not contain "..", and does not end with .co
    # except for .edu.co, .org.co, .com.co, .delete.com, or gentherm.com
    if re.match(pattern, email) and ".." not in email and not (email.endswith(".co") and not (email.endswith((".edu.co", ".org.co", ".com.co")))):
        return True
    return False

# Check all rows in column A
try:
    for row in range(1, sheet.max_row + 1):
        cell = f"A{row}"
        email = sheet[cell].value
        if email:  # Check that the cell is not empty
            validity = "Valid" if is_email_valid(email) else "Invalid"
            sheet[f"B{row}"] = validity  # Write the result in column B
        else:
            sheet[f"B{row}"] = "Invalid"  # Mark as invalid if the cell is empty

    # Determine the output filename
    base_filename = "result_verification"
    extension = ".xlsx"
    filename = base_filename + extension
    counter = 1

    # Increment the filename if it already exists
    while os.path.exists(filename):
        filename = f"{base_filename}_{counter}{extension}"
        counter += 1

    # Save the results in a new file
    wb.save(filename)
    print("The script has been executed successfully.")
    print(f'Your file has been saved as "{filename}".')

except Exception as e:
    with open("error_log.txt", "w") as error_file:
        error_file.write(f"An error occurred during the script execution: {e}\n")
    print(f"An error occurred during the script execution: {e}. Details saved in error_log.txt.")
    input("Press Enter to close the window.")

# Keep the window open after execution
input("Press Enter to close the window.")
