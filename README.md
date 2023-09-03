# autmated_report_generation
This app takes in an excel spreadsheet and word template and is able to generate reports from the excel data to the word template

Check "requirements.txt" file for the required Python packages.
The word template is included. Names within double curly braces are considered as variables. 
The PDF/Word report generated will be based on the .dotx word template format.
How this works:
1)Reads values from an excel sheet. The values that need to be readed are obtained from the hard coded column numbers, so adjust the column numbers in the code to fit the required data.
2)Accesses the .dotx MS Word Template file and reads the variables in it.
3)Assigns the values from the Excel sheet to the Variables in the dotx Template.
4)Saves the report as a word document and repeats for the next row of the excel sheet.
5)Once all the rows are read, the loop terminates.
6)If you need PDF reports, uncomment the line relevant to that. Then all the word reports will be converted to PDF.
