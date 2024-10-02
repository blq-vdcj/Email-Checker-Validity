# Procedure to Run the Script on Any Machine

## 1. Install Python
1. **Download Python**:
   - Visit the official Python website: [https://www.python.org/downloads/](https://www.python.org/downloads/).
   - Download the version compatible with your operating system (Windows, macOS, or Linux).

2. **Install Python**:
   - Run the downloaded installer.
   - **Check the box** that says "Add Python to PATH" during installation. This will make it easier to use Python from the terminal.
   - Follow the installation steps until complete.

## 2. Install Required Library (openpyxl)
1. **Open a Terminal**:
   - On Windows: Press `Win + R`, type `cmd`, and press `Enter`.
   - On macOS/Linux: Open the "Terminal" application.

2. **Install the openpyxl Library**:
   - Type the following command in the terminal and press `Enter`:
     ```sh
     pip install openpyxl
     ```
   - If you have multiple versions of Python installed, use `pip3` instead of `pip`:
     ```sh
     pip3 install openpyxl
     ```

## 3. Prepare the Excel File
1. **Create the Excel File**:
   - Create an Excel file named `email.xlsx`.
   - Make sure that the email addresses to be verified are placed in column A, starting from cell `A1`.

2. **Save the File**:
   - Save this file in the same directory where you will place the Python script for easy access.

## 4. Download and Prepare the Python Script
1. **Create the Python Script**:
   - Copy the provided Python code into a new text editor (such as **Notepad++**, **Visual Studio Code**, or even **Notepad**).
   - Save the file with a `.py` extension, e.g., `emailchecker.py`.

2. **Check the Directory**:
   - Place `email.xlsx` and `emailchecker.py` in the same directory (e.g., on the Desktop or in a dedicated folder).

## 5. Run the Script
1. **Open Terminal**:
   - Navigate to the directory where the script and the Excel file are located.
   - For example, if the files are on the Desktop:
     ```sh
     cd Desktop
     ```
   
2. **Run the Script**:
   - Type the following command in the terminal:
     ```sh
     python emailchecker.py
     ```
   - If you have multiple versions of Python installed, use `python3` instead of `python`:
     ```sh
     python3 emailchecker.py
     ```

## 6. Check the Results
1. **Result File**:
   - After execution, a file named `result_verification.xlsx` (or `result_verification_1.xlsx`, etc. if there are multiple versions) will be created in the same directory.
   - Open this file to see the results, which will be in column B with "Valid" or "Invalid" for each email address.

## 7. Error Handling
- If an error occurs while executing the script, a file named `error_log.txt` will be generated with a description of the error.
- Make sure to check this file to understand what went wrong.

## Summary of Key Commands
- **Install Python**: [https://www.python.org/downloads/](https://www.python.org/downloads/)
- **Install openpyxl**:
  ```sh
  pip install openpyxl
