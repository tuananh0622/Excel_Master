Attribute VB_Name = "Module1"
Sub Get_Data_From_File()
    Dim FileToOpen1 As Variant
    Dim FileToOpen2 As Variant
    Dim OpenBook1 As Workbook
    Dim OpenBook2 As Workbook
    Application.ScreenUpdating = False
    FileToOpen1 = Application.GetOpenFilename(Title:="Browse for your file & Import range", FileFilter:="Excel Files:=(*.xls*),*xls*")
    If FileToOpen1 <> False Then
        Set OpenBook1 = Application.Workbooks.Open(FileToOpen1)
        OpenBook1.Sheets(1).Range("A3:G1442").Copy
        ThisWorkbook.Worksheets("1").Range("A3").PasteSpecial xlPasteValuesAndNumberFormats
        ThisWorkbook.Save
        OpenBook1.Close False
    End If
    FileToOpen2 = Application.GetOpenFilename(Title:="Browse for your file & Import range", FileFilter:="Excel Files:=(*.xls*),*xls*")
    If FileToOpen2 <> False Then
        Set OpenBook2 = Application.Workbooks.Open(FileToOpen2)
        OpenBook2.Sheets(1).Range("A3:E146").Copy
        ThisWorkbook.Worksheets("1").Range("I3").PasteSpecial xlPasteValuesAndNumberFormats
        ThisWorkbook.Save
        OpenBook2.Close False
    End If
    Application.ScreenUpdating = True
End Sub
