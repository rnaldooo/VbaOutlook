Attribute VB_Name = "M�dulo1"
Sub BUandSave2()
'Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
'Saves the current file to a backup folder and the default folder
'Note that any backup is overwritten
Dim MyDate
MyDate = Date    ' MyDate contains the current system date.
Dim MyTime
MyTime = Time    ' Return current system time.
Dim TestStr As String
TestStr = Format(MyTime, "hh.mm.ss")
Dim Test1Str As String
Test1Str = Format(MyDate, "DD-MM-YYYY")

Application.DisplayAlerts = False
Dim ttt As String
ttt = "C:\Users\aaa\Desktop\teste" & Test1Str & " " & TestStr & " " & ActiveWorkbook.Name
'SaveWorkbookAsNewFile (ttt)
'
'Application.Run ("SaveFile")
'
ActiveWorkbook.SaveCopyAs FileName:="C:\Users\aaaa\Desktop\teste" & Test1Str & " " & TestStr & " " & ActiveWorkbook.Name
ActiveWorkbook.Save
Application.DisplayAlerts = True
End Sub


Private Sub SaveWorkbookAsNewFile(NewFileName As String)
    Dim ActSheet As Worksheet
    Dim ActBook As Workbook
    Dim CurrentFile As String
    Dim NewFileType As String
    Dim NewFile As String
 
    Application.ScreenUpdating = False    ' Prevents screen refreshing.

    CurrentFile = ThisWorkbook.FullName
 
    NewFileType = "Excel Files 1997-2003 (*.xls), *.xls," & _
               "Excel Files 2007 (*.xlsx), *.xlsx," & _
               "All files (*.*), *.*"
 
    NewFile = Application.GetSaveAsFilename( _
        InitialFileName:=NewFileName, _
        fileFilter:=NewFileType)
 
    If NewFile <> "" And NewFile <> "False" Then
        ActiveWorkbook.SaveAs FileName:=NewFile, _
            FileFormat:=xlNormal, _
            Password:="", _
            WriteResPassword:="", _
            ReadOnlyRecommended:=False, _
            CreateBackup:=False
 
        Set ActBook = ActiveWorkbook
        Workbooks.Open CurrentFile
        ActBook.Close
    End If
 
    Application.ScreenUpdating = True
End Sub


Sub SaveWithoutCode()
    Dim ws As Worksheet
    Dim i As Integer
    Dim sarrWS() As String
    ReDim sarrWS(1 To ThisWorkbook.Worksheets.Count)
    i = 0
    For Each ws In ThisWorkbook.Worksheets
        i = i + 1
        sarrWS(i) = ws.Name
    Next ws
    ThisWorkbook.Worksheets(sarrWS()).Copy
End Sub






Sub Macrrdfo4()
'
' Macro4 Macro
'

'

    Sheets("Ident. Amostras").Select
    Sheets("Ident. Amostras").Copy Before:=Workbooks("Pasta1").Sheets(1)
End Sub
