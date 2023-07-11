# Excel_Master
If you want to convert multiple xls files to xlsx files at once without saving one by one, here, I will talk about a VBA code for you, please do with the following steps:

1. Hold down the ALT + F11 keys to open the Microsoft Visual Basic for Applications window.

2. Click Insert > Module, and paste the following code in the Module Window.
_Code is below:_

Sub ConvertToXlsx()
'Updateby Extendoffice
Dim strPath As String
Dim strFile As String
Dim xWbk As Workbook
Dim xSFD, xRFD As FileDialog
Dim xSPath As String
Dim xRPath As String
Set xSFD = Application.FileDialog(msoFileDialogFolderPicker)
With xSFD
.Title = "Please select the folder contains the xls files:"
.InitialFileName = "C:\"
End With
If xSFD.Show <> -1 Then Exit Sub
xSPath = xSFD.SelectedItems.Item(1)
Set xRFD = Application.FileDialog(msoFileDialogFolderPicker)
With xRFD
.Title = "Please select a folder for outputting the new files:"
.InitialFileName = "C:\"
End With
If xRFD.Show <> -1 Then Exit Sub
xRPath = xRFD.SelectedItems.Item(1) & "\"
strPath = xSPath & "\"
strFile = Dir(strPath & "*.xls")
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Do While strFile <> ""
If Right(strFile, 3) = "xls" Then
Set xWbk = Workbooks.Open(Filename:=strPath & strFile)
xWbk.SaveAs Filename:=xRPath & strFile & "x", _
FileFormat:=xlOpenXMLWorkbook
xWbk.Close SaveChanges:=False
End If
strFile = Dir
Loop
Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub

4. Then, press F5 key to run this code, and a window will be displayed, please select a folder which contains the xls files that you want to convert
5. Then,click OK, another window is popped out, please select a folder path where you want to output the new converted files
6. And then, clik OK, after finishing the conversion, you can go to the specified folder to preview the converted result, see screenshots:
