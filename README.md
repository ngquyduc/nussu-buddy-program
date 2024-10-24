# Buddy Matching Script

## Overview
The Buddy Matching Script is a powerful tool designed to efficiently pair over 700 participants in the NUSSU Buddy Program. By leveraging Python and pandas, this script simplifies the buddy matching process, ensuring a smooth and effective experience for everyone involved.

## Prerequisites
- **Python:** Ensure Python is installed on your machine. Check with `python --version` in the terminal.
- **Pandas:** Install pandas with the command: `pip install pandas`.
- **Executable Script:** Make the script executable if necessary by running `chmod +x script.py` in the terminal.

## Important Notes
- **Input File:** Rename your input file to `input.xlsx` for seamless processing.
- **Column Names:** Ensure that the column names in your input file match those specified in the script. If you modify the input format, remember to update the script accordingly.
- **Clean Input Data:** Remove any unnecessary rows at the beginning of the input file so that the first row contains the column headers.

## How to Run
1. Open the project folder in VSCode or navigate to it using the terminal.
2. Execute the script by running: 
   ```bash
   python3 script.py
   ```
   or
   ```bash
   python script.py
   ```
3. The output will be saved in `output.xlsx` within the same directory.