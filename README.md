# Sorting_LinkIDs_Arcadis

This project involves a script to reorder link IDs in an Excel file based on specific criteria resembling a Linked List. The goal is to ensure that each link ID in a row follows a sequence where the last three digits of one link match the first three digits of the next link. This project is designed to maintain data integrity and improve readability in large datasets.

## Table of Contents
- [Installation](#installation)
- [Usage](#usage)
- [How It Works](#how-it-works)
- [Algorithm](#algorithm)
- [Files](#files)

## Installation
To set up the environment and install the necessary dependencies, follow these steps:

### Create a Virtual Environment
### Activate Virtual Environment
### pip install -r requirments.txt


# Usage

## Rename Your Excel File
Ensure your Excel file is named `original.xlsx` and placed in the same directory as the script.

## Run the Script
Execute the script to start the reordering process.

## Check Output
The script will generate two output files:
- `ordered.xlsx`: The updated Excel file with reordered link IDs.
- `change_log_fixed.txt`: A text file logging all changes made during the process.

# How It Works
The script performs the following steps:
1. **Load the Excel File**: The script loads the specified Excel file and reads the target sheet into a DataFrame.
2. **Extract Link ID Columns**: It extracts columns that contain 'Link IDs' for processing.
3. **Reorder Link IDs**: A function reorders the link IDs in each row. It starts from the head link and iteratively finds the next link by matching the last three digits of the current link with the first three digits of the potential next link.
4. **Update DataFrame**: The script updates the DataFrame with the newly ordered links while preserving the original structure.
5. **Save Results**: The updated DataFrame is saved to a new Excel file, and a change log is generated to record all modifications.

# Algorithm
The algorithm to reorder the link IDs follows these steps:
1. **Initialize the Dictionary**: Create a dictionary to map the start and end points of valid links.
2. **Create the Ordered List**: Start with the head link and iteratively find the next link by matching the last three digits of the current link with the first three digits of the potential next link. Append each matched link to the ordered list.
3. **Update DataFrame**: Replace the original links in the DataFrame with the ordered links.
4. **Logging Changes**: Log all changes made to ensure transparency and traceability.

# Files
- `Script.py`: The main script to reorder the link IDs.
- `original.xlsx`: The original Excel file to be processed.
- `ordered.xlsx`: The updated Excel file with reordered link IDs.
- `change_log_fixed.txt`: A text file logging all changes made during the process.