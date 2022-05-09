Attribute VB_Name = "Módulo5"
Sub tsdftt()
'Make the excel file that runs the software the active workbook
ThisWorkbook.Activate

'The first sheet used as a temporary place to hold the data
ThisWorkbook.Worksheets(1).Cells.Copy

'Create a new Excel workbook
Dim NewCaseFile As Workbook
Dim strFileName As String

Set NewCaseFile = Workbooks.Add
With NewCaseFile
    Sheets(1).Select
    Cells(1, 1).Select
End With

ActiveSheet.Paste
End Sub


Sub dsafsaf()
Dim wb As Workbook
 
Worksheets("Ident. Amostras").Copy
Set wb = ActiveWorkbook
wb.SaveAs "New Report.xls"
wb.Close

Worksheets("Ident. Amostras").Copy
Set wb = ActiveWorkbook
wb.SaveAs "New Report2.xls"
wb.Close


End Sub
