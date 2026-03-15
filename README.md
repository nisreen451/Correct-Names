This Python script processes a detailed payments Excel file, cleans and standardizes the Beneficiary Name English column, and saves the corrected file for further use or reporting.

Functionality Overview

Read input file

Loads the Excel file located at:
DetailedPayment_for_beneficaries_52202682436.xlsx

Name corrections

Uses a dictionary name_corrections to standardize common variations, typos, or abbreviations in names.

Examples:

"IBRM" → "IBRAHIM"

"MAHM" / "Mahd" → "MAHMOUD"

"A" or "A." followed by a token → merged as "ABD AL <next token>"

Clean function

Splits each name into tokens.

Applies replacements from the dictionary.

Handles special cases for “A” / “A.” merges.

Rejoins tokens to form a cleaned full name.

Load workbook

Opens the Excel workbook using openpyxl.

Selects only the sheet named Payments.

Scans the first 20 rows to locate the header row for Beneficiary Name English.

Update column values

Loops through each cell in the Beneficiary Name English column.

Applies the clean_name() function to standardize the name.

Save output

Saves the updated workbook in the OutputFiles folder.

Filename: payment summary with corrected names.xlsx

Preserves the original Excel file extension.

Output confirmation

Prints the full path of the saved file after completion.

How to Use (Read Me)

Place the input file in the InputFiles folder.

Update the name_corrections dictionary if new name variations are found.

Ensure the OutputFiles folder exists.

Run the script.

Check the OutputFiles folder for a cleaned version of the payments sheet.

Before / After Example

Original Name	Cleaned Name
IBRM MAHM A HLA	IBRAHIM MAHMOUD ABD ALHALA
Mohd CHARIF	MOHAMMAD SHARIF
CHIHAB MUST	SHIHAB MUSTAFA
