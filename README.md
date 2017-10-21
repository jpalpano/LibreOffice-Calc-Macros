# LibreOffice-Calc-Macros
Useful BASIC macro codes for LibreOffice Calc

The custom functions and procedure aims to help new LibreOffice Calc BASIC programmers who are coming rom Microsoft Excel VBA background.

The codes given aims to, somehow, simulate the common functions and procedures in Excel such as:

ACCESSING SHEETS, RANGES, AND COLUMNS
1) Sheets("Sheet1")
2) Range("A1")
3) Columns("A")

SETTING FOCUS TO SHEETS OR RANGES
1) Sheets("Sheet1").Activate
2) Range("A1").Activate

COPYING VALUES TO CLIPBOARD
*) Range("A1").Copy

ACCESSING THE LAST NON-EMPTY ROW
*) Range("A" & Rows.Count).End(xlUp).Row


The other useful procedures are:
1) Removing Duplicates
2) Applying a Standard Filter
3) Checking if the given string is an email address
4) Comparing a given string with a Regular Expression
