# Reports Package

## Overview
The `Reports Package` automates the process of generating study visit reports using Python. Traditionally, this process was manual and time-consuming, requiring multiple days and people to verify and insert data. This package streamlines report creation by extracting data from REDCap fields (VABS, MSEL, and PLS), generating structured tables and descriptions, and inserting names and signatures automatically.

### Author
**Michael Khela**  
Email: [michael.khela@childrens.harvard.edu](mailto:michael.khela99@gmail.com)

### Supervisor
**Carol Wilkinson**

### Contributors
- McKena Geiger  
- Sophie Hurewitz  
- Meagan Tsou 

## Requirements
- **Python** (recommended version: 3.12.1)
- Required Python libraries:
  - `pandas`
  - `python-docx`

To install dependencies, run:
```sh
pip install pandas python-docx
```

## Installation
1. Clone or download the `Reports Package` repository.
2. Copy the package folder to your working directory.
3. Ensure you have Python and Anaconda installed (Spyder is recommended for running the script).

## Usage
### Setting Up
1. Open **Spyder** via Anaconda Navigator.
2. Open `Run_Report_Automation.py` in Spyder.
3. Fill in necessary inputs, such as:
   - `subject ID`
   - `administrators`
   - `file paths`
   - `REDCap export filename`
4. Ensure all input files are closed before running the script.

### Running the Script
Execute the script by pressing the green **Run** button in Spyder.
```sh
python Run_Report_Automation.py
```
### Output
- The script generates a study visit report in the `Created/` directory.
- Names, dates, and signatures are automatically inserted.

## Troubleshooting
- Ensure **input files are closed** before running the script.
- Install missing packages if prompted (`pip install package-name`).
- If an error occurs:
  1. Verify your file paths and filenames.
  2. Ensure your REDCap export file is in CSV format.
  3. Contact Michael Khela with a screenshot of the error.

## Contact
For troubleshooting and support, contact **Michael Khela** at [michael.khela@childrens.harvard.edu](mailto:michael.khela99@gmail.com).

