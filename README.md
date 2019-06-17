# HW.Unit2
Homework Unit 2 | Assignment - The VBA of Wall Street

## Files Uploaded:
1. 1_OptionBox.frm
2. 1_OptionBox.frx
3. 2_Module1_Easy.bas
4. 3_Module2_Moderate.bas
5. 4_Module3_Hard.bas
6. 5_Module4_Challenge.bas
7. DanielMJones.HW.Unit2.UserForm.ScreenShot.pdf
8. HW.Unit2.Screen.Shots

## VBA Process Description:
The Modules 1, 2, 3, and 4 are associated with the HW Unit 2 difficulty levels. The VBA code performs the following actions:
  1. Easy (Module 1):
      a. Displays each unique ticker symbol
      b. Calculates the total stock volume for each ticker symbol
  2. Moderate (Module 2):
      a.  Calculates Yearly Change from opening price at the beginning of a given year to the closing price at the end of that year for each unique ticker symbol
      b.  Calculates Percent Change from opening price at the beginning of a given year to the closing price at the end of that year for each unique ticker symbol
      c. Processes conditional formatting that will highlight positive change in Yearly Change in green and negative change in red for each unique ticker symbol
  3. Hard (Module 3):
      a. Calculates the stock with the "Greatest % increase", "Greatest % Decrease" and "Greatest total volume
  4. Challenge (Module 4):
      a. Runs script to process the all modules on every worksheet in workbook
      
## Note:
  For the Multiple_year_stock_data file, the VBA code takes approximately 10 minutes to run for each year, and approximately 30 minutes for the whole file.  A message box appears once the VBA code has completed running for each worksheet; if the VBA code to run through the whole file is ran, the user will need to click the message box once the VBA code has run for a particular worksheet before it moves to the next worksheet.
  
## UserForm (OptionBox):
  An optionbox opens up when the Excel file opens, that allows the user to make the following choices:
    1. Select the worksheets the user wants to run the VBA code on
         > Left box with worksheets from workbox
         > Click the "Add" button to add to the right box
    2. Remove any worksheets the user selected to run, but decides not to run
         > Click the "Remove" button to remove from the right box
    3. Process (run the VBA code) the selected worksheets
        > Click the "Process Selected" button
    4. Reset, or delete the results of the VBA code, the current worksheet the user is on
        > Click the "Reset Current Sheet" button
    5. Reset all worksheets in the workboox
        > Click the "Reset All Sheets" button
    6. Process (run the VBA code) on all (every) worksheets in the workbook
        > Click the "Process ALL Sheets" button
    7. Exit out of the OptionBox selections
        > Click the "Exit" button
      
