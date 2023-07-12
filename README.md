# Excel_Master
If you want to convert multiple xls files to xlsx files at once without saving one by one, here, I will talk about a VBA code for you, please do with the following steps:
1. Hold down the ALT + F11 keys to open the Microsoft Visual Basic for Applications window.
2. Click Insert > Module, and paste the excel_xls_to_xlsx.bas code in the Module Window.
4. Then, press F5 key to run this code, and a window will be displayed, please select a folder which contains the xls files that you want to convert
5. Then,click OK, another window is popped out, please select a folder path where you want to output the new converted files
6. And then, clik OK, after finishing the conversion, you can go to the specified folder to preview the converted result, see screenshots:

# Get data from files
Insert Get_data_from_file.bas into your VBA code
Use this code line to ignore excel warning about cut or copy a large amount of data:
ThisWorkbook.Save
Choose the range of your slave workbook by using this line of code:
OpenBook1.Sheets(1).Range("A3:G1442").Copy
Choose the sheet and where to put it down in your master workbook by using this line of code:
ThisWorkbook.Worksheets("1").Range("A3").PasteSpecial xlPasteValuesAndNumberFormats

# Jump to first and last sheet
Insert Jump_to_1st_last.bas into your VBA code
Click "Macros", then click "Option" and set the hotkey you want to use (Crtl+ letter)
