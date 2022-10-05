# Buddy Temperature Data Analyzer

## Purpose:

The purpose of this program is to quickly analyze large numbers of test data files. Each file is analyzed and a summary of all data files is tabulated and saved to an excel file.

## Program instructions

All input and output files must be closed before running program so that they can be edited. An error will occur if the files cannot be written.

### Input File Selection

The user is prompted to select one or more files to analyze via a file selection box.

The file must be a valid excel file with extension ".xls" or ".xlsx"

The flowrate, input temperature target, battery id# and disposable id# must be recorded in the data set's filename. The filename(s) must be in the following format: 
>{flowrate}ml_m {input temp target}C batt{battery id#} disp{disposable id#} unit{device id#}.xlsx

optionally, the device id# can be ommitted

optionally, additional text can be added to the end of the filename if desired
>
>e.g. "92ml_m 10C battA4 disp6 unit4.xlsx"
>
>or "102ml_m 20C battA1 disp3 environmental chamber test.xlsx"

{flowrate} and {input temp target} must be numbers.

### Output File Selection

The user is prompted to select an excel output file to append data summary via a file selection box.

The output file can contain previously compiled data or no data.

Calculations are performed and appended to the output file. If no data summary table exists in the output file, it will be created with headers.

Calculations are also appended to each original data file in a new worksheet.

## Description of Algorithm

This program is intended to be converted to a .exe file so that it can easily be run on a lab computer. For this reason, all of the error messages are presented as a window popup using the tkinter library.

Required external libraries: 
- pandas: Converts excel files to a database that can be manipulated using pandas
- openpyxl: Used to write and format excel files

When the program is launched the user is prompted with a file open dialogue box for the input files and then a second one for the output file. If the files are not excel files, the program shows an error dialogue box and ends the program. 

The output file selected is opened using openpyxl. If the user has this file open, the program will close with an error message because it cannot be written to. If the file does not contain an existing data summary table, a new table is created with headers.

The information in the table includes:

- **Date**: The date of test.
- **Disposable**: Disposable id# used during test. Read from filename.
- **Battery**: Battery id# used during test. Read from filename.
- **Input Temp Target (°C)**: Target input fluid temperature. Read from filename. This is not used in any calculation, it is intended only to help with sorting similar tests in excel.
- **Flowrate (mL/min)**: Measured flowrate during test. Read from filename.
- **Unit**: Device id#. Read from filename.
- **Input Temp Mean (°C)**: Calculated mean input temperature
- **Steady-State Output Temp (°C)**: Calculated mean output temperature. Only calculates the average of data during minutes 5 - 12 because the data is typically steady during this period.
- **Reservoir Temp Mean (°C)**: Optional. Calculates the mean fluid reservoir temperature.
- **Startup Time**: Time that the output fluid temperature took to reach the minimum specification (36°C).
- **Peak Temp (°C)**: Peak output temperature during the test.
- **Test Time > 36°C**: Total time the test delivered fluid above the minimum specification (36°C). Checks for the time the output temperature drops below 36°C for 10 consecutive seconds.
- **Fluid Infused (mL)**: Total fluid infused above the minimum specification. Calculated from test time and flowrate.
- **Battery Time**: Length of time the battery charge lasted. 
- **ΔT x Time x Flowrate**: Calculation used to compare across multiple tests.
- **Filename**: The name of the original file for reference.
- **Comments**: This column is left blank by the program in case the user needs to add comments later.

The main loop of the program loops through all of the input files selected. One at a time, they are read using pandas library, and calculations are performed.

The data summary is appended to the summary table. Additionally the data summary is added to the original data file for reference in a new sheet.
