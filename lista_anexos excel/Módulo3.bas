Attribute VB_Name = "Módulo3"
Dim iRow As Integer

Sub PegaOriginal()
Call SelectFolder(3)
End Sub

Sub SelectFolder(ilinha As Integer)
    Dim diaFolder As FileDialog
    Set diaFolder = Application.FileDialog(msoFileDialogFolderPicker)
    diaFolder.AllowMultiSelect = False
    diaFolder.Show
    ActiveSheet.Range("E" & ilinha & "").Value = diaFolder.SelectedItems(1) 'diaFolder.InitialFileName
    'ActiveSheet.Range("F" & ilinha & "").Value = Replace(diaFolder.SelectedItems(1), diaFolder.InitialFileName, "")
    Set diaFolder = Nothing
  
End Sub


Sub ListFolders()
 iRow = 15
 Sheets("inicio").Select
 Range("a1").Select
 Call ListMyFiles(Range("e3").Value, Range("e4").Value)
 End Sub
 
 'Look at the new file in the VBE menu Tools -> References... whether Microsoft Scripting Runtime is checked.
Sub ListMyFolders(mySourcePath As String, IncludeSubfolders As Boolean)
 Set MyObject = New Scripting.FileSystemObject
 Set mySource = MyObject.GetFolder(mySourcePath)
 On Error Resume Next
 For Each myFolder In mySource.Files 'SubFolder
 iCol = 2
 Cells(iRow, iCol).Value = myFolder.Path
 iCol = iCol + 1
 Cells(iRow, iCol).Value = myFolder.Name
 iCol = iCol + 1
 Cells(iRow, iCol).Value = myFolder.DateLastModified
 iRow = iRow + 1
 Next
 If IncludeSubfolders Then
 For Each mySubFolder In mySource.SubFolders
 Call ListMyFolders(mySubFolder.Path, True)
 Next
 End If
 End Sub


Sub ListMyFiles(mySourcePath As String, IncludeSubfolders As Boolean)
    Set MyObject = New Scripting.FileSystemObject
    Set mySource = MyObject.GetFolder(mySourcePath)
    On Error Resume Next
     For Each MyFile In mySource.Files
       If InStr(MyFile.Path, Sheets("inicio").Range("E5").Value) <> 0 Then
            iCol = 5
            Cells(iRow, iCol).Value = MyFile.Path
            iCol = iCol + 1
            Cells(iRow, iCol).Value = Replace(MyFile.Name, " ", "")
            iCol = iCol + 1
            Cells(iRow, iCol).Value = MyFile.Size
            iCol = iCol + 1
            Cells(iRow, iCol).Value = MyFile.DateLastModified
            iRow = iRow + 1
            End If
         Next
    
    Columns("C:E").AutoFit
    If IncludeSubfolders Then
        For Each mySubFolder In mySource.SubFolders
            Call ListMyFiles(mySubFolder.Path, True)
        Next
    End If
End Sub


Sub limparcell()
'
' Macro1 Macro
'

'
    Range("E15:H305").Select
    Selection.ClearContents
    Range("A1").Select
End Sub
