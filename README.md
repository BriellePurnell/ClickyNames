# ClickyNames

Clicky Names is an application that builds a table, using the pandastable module, from a given Excel file. 
It is used to easily parse full names into their last, first, and middle names by double clicking (instead
of cut and paste). This tool is meant to work with a specific structure of spreadsheet and to solve a specific employee problem.

## Installation

Download the ClickyNames.exe file and run it. (Only the employee has the .exe file)

## Usage

The names you want to parse must be in the second column of the Excel sheet that you are importing. Clicky Names 
inserts three empty columns after this second column. Those three columns are Last Name, First Name, and 
Middle Name.

When you first get Clicky Names running, you first must select an Excel file by choosing File > Open in the menu
bar. Now that the file has been imported, you can click on a full name and a pop up window will appear. Double
click on a last name first which will transfer it to the Last Name column, then First Name, then Middle Name
or initial. Left clicking on the other columns in the pop up window will clear the First, Last, and Middle Name
columns. Right clicking anywhere will "save" your changes in the popup window to the table. After all your names
have been parsed, you can save the table to another Excel file by choosing File > Save in the menu bar.

The saved Excel file will have a column that shouldn't be there. It's the first column and it represents the row
index from the table in Clicky Names. To delete it, just right click on the column header in Excel and select
Delete.

## Author

Brielle Purnell
