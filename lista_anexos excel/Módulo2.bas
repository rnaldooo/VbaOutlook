Attribute VB_Name = "Módulo2"

Sub CopiandoPARA()
Dim objExcel1 As Object
Dim objExcel2 As Object


' EXCEL - 1 ************************ abrindo arquivo - referencia necessaria: ole automation
Set objExcel1 = GetObject(, "Excel.Application")                ' define excel como o objeto

' EXCEL - 2 ************************ abrindo arquivo - referencia necessaria: ole automation
Set objExcel2 = CreateObject("Excel.Application")
objExcel2.Workbooks.Add
objExcel2.Visible = True

        objExcel2.Columns("A:A").ColumnWidth = 24
        objExcel2.Columns("B:B").ColumnWidth = 10
        objExcel2.Columns("C:C").ColumnWidth = 19
        objExcel2.Columns("D:D").ColumnWidth = 10
        objExcel2.Columns("E:E").ColumnWidth = 13
        objExcel2.Columns("F:F").ColumnWidth = 13
        objExcel2.Columns("G:G").ColumnWidth = 16

objExcel1.Sheets("Ident. Amostras").Select
objExcel1.Range("A1").Select
    objExcel1.ActiveSheet.Shapes.Range(Array("Picture 5")).Select
    objExcel1.Application.CutCopyMode = False
    objExcel1.Selection.Copy
        objExcel2.Sheets("Plan1").Select
        objExcel2.Range("A1").Select
        AppActivate objExcel2
        objExcel2.ActiveSheet.Paste
    objExcel1.Sheets("Ident. Amostras").Select
    objExcel1.ActiveSheet.Shapes.Range(Array("Picture 8")).Select
    objExcel1.Selection.Copy
        objExcel2.Sheets("Plan1").Select
        objExcel2.Range("G1").Select
        objExcel2.ActiveSheet.Paste
        objExcel2.Range("A1").Select


    
        


'objExcel1.Sheets("PlanAutocad").Select                      '      seleciona PlanAutocad
'objExcel1.Sheets("PlanAutocad").Cells.Clear                 '          limpa PlanAutocad
'Excel.Visible = True
'objExcel1.Visible = False                                   '               oculta excel
'objExcel1.Range("A1").Select                                '               seleciona A1
'objExcel1.Cells(1, 1).Value = "reinaldo dimensiona excel-cad" '            coloca titulo

End Sub




Sub simples()
Dim objExcel1 As Object
Dim objExcel2 As Object


' EXCEL - 1 ************************ abrindo arquivo - referencia necessaria: ole automation
Set objExcel1 = GetObject(, "Excel.Application")                ' define excel como o objeto

' EXCEL - 2 ************************ abrindo arquivo - referencia necessaria: ole automation
Set objExcel2 = CreateObject("Excel.Application")
objExcel2.Workbooks.Add
objExcel2.Visible = True

    objExcel1.Sheets("Ident. Amostras").Select

    'Replace "Sheet1" with the name of the sheet to be copied.
    objExcel1.Sheets("Ident. Amostras").Copy Before:=objExcel2.ActiveWorkbook.Sheets(1)

   ' objExcel1.Sheets("Ident. Amostras").Copy ' Before:=Workbooks("Pasta1").Sheets(1)
End Sub
   
   
   
   
   
   Sub copyThem()
    Dim ws As Worksheet, ss As Worksheet, FolderName As String, wb As Workbook
    Application.ScreenUpdating = False
    FolderName = ThisWorkbook.Path
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Temp.sheet" And ws.Name <> "inicio" Then
            If wb Is Nothing Then
                ws.Copy
                Set wb = ActiveWorkbook
            Else
                ws.Copy after:=ss
            End If
            Set ss = ActiveSheet
        End If
    Next ws
    Range("i9").Value = 1
    ThisWorkbook.Activate
    'Wb.Sheets(1).Select
    wb.SaveAs FolderName & "\" & wb.Sheets(1).Name & ".xlsx"
    wb.Close False
    'MsgBox "Look in folder " & FolderName & " for files ..."
    Application.ScreenUpdating = True
End Sub



 
   Sub geraaarq()
   Dim ia                As Integer
   Dim itotal            As Integer
   Dim socaminho         As String
   Dim sdcaminho         As String
   Dim snarquivo         As String
   Dim spcaminho         As String
   socaminho = ThisWorkbook.Sheets("inicio").Range("e3").Value
   sdcaminho = ThisWorkbook.Sheets("inicio").Range("e12").Value
   itotal = ThisWorkbook.Sheets("inicio").Range("e14").Value
   
   For ia = 1 To itotal
   snarquivo = ThisWorkbook.Sheets("inicio").Range("f" & ia + 14 & "").Value
   snarquivo = Replace(snarquivo, ".xls", "")
   Dim ws As Worksheet, ss As Worksheet, FolderName As String, wb As Workbook
    Application.ScreenUpdating = False
    'FolderName = ThisWorkbook.Path
    
    Worksheets("Ident. Amostras").Copy
    Set wb = ActiveWorkbook
    
    Range("a5").Value = ThisWorkbook.Sheets("inicio").Range("I" & ia + 14 & "").Value 'OS
    Range("a6").Value = ThisWorkbook.Sheets("inicio").Range("K" & ia + 14 & "").Value 'DATA (-ALGUNS DIAS)
    Range("B5").Value = ThisWorkbook.Sheets("inicio").Range("J8").Value 'OBRA
    Range("B6").Value = ThisWorkbook.Sheets("inicio").Range("J9").Value 'CLIENTE
    Range("E5").Value = ThisWorkbook.Sheets("inicio").Range("J" & ia + 14 & "").Value 'LOTE GRD
    Range("E6").Value = ThisWorkbook.Sheets("inicio").Range("J10").Value 'INSPETOR
    
    Range("E10").Value = ThisWorkbook.Sheets("inicio").Range("L" & ia + 14 & "").Value 'DATA LIMITE(GRD)
    Range("F10").Value = ThisWorkbook.Sheets("inicio").Range("e10").Value 'RESPONSÁVEL
    Range("E26").Value = ThisWorkbook.Sheets("inicio").Range("L" & ia + 14 & "").Value 'DATA LIMITE(GRD)
    Range("F26").Value = ThisWorkbook.Sheets("inicio").Range("e10").Value 'RESPONSÁVEL
    
   
'    For Each ws In ThisWorkbook.Worksheets
'        If ws.Name <> "Temp.sheet" And ws.Name <> "inicio" Then
'            If wb Is Nothing Then
'                ws.Copy
'                Set wb = ActiveWorkbook
'            Else
'                ws.Copy after:=ss
'            End If
'            Set ss = ActiveSheet
'        End If
'    Next ws
'   Range("i9").Value = 3
    spcaminho = sdcaminho & "\checklist_" & snarquivo & ".xlsx"
    ThisWorkbook.Activate
    'Wb.Sheets(1).Select
   ' Wb.SaveAs FolderName & "\" & Wb.Sheets(1).Name & ".xlsx"
   If Len(Dir(spcaminho)) > 0 Then
            DeleteFile (spcaminho)
   End If
     wb.SaveAs spcaminho
    wb.Close False
    
    'MsgBox "Look in folder " & FolderName & " for files ..."
    Application.ScreenUpdating = True
   
   Next
   
End Sub

Sub DeleteFile(ByVal FileToDelete As String)
   If FileExists(FileToDelete) Then 'See above
      SetAttr FileToDelete, vbNormal
      Kill FileToDelete
   End If
End Sub

Function FileExists(ByVal FileToTest As String) As Boolean
   FileExists = (Dir(FileToTest) <> "")
End Function

Sub deonde()

    Prompt = "Entre com o caminho onde quer salvar"
    Title = "Pasta de Destino"
    StrSavePath = BrowseForFolder(&H200)
'Const BIF_NEWDIALOGSTYLE = &H40
'Const BIF_NONEWFOLDERBUTTON = &H200
'Const BIF_RETURNONLYFSDIRS = &H1
  
    If StrSavePath = "" Then
GoTo ExitSub:
    
    End If
    ThisWorkbook.Sheets("inicio").Range("e3").Value = StrSavePath
ExitSub:
End Sub

Sub paraonde()

    Prompt = "Entre com o caminho onde quer salvar"
    Title = "Pasta de Destino"
    StrSavePath = BrowseForFolder(0)
' 0 new folder
'Const BIF_NEWDIALOGSTYLE = &H40
'Const BIF_NONEWFOLDERBUTTON = &H200
'Const BIF_RETURNONLYFSDIRS = &H1
  
    If StrSavePath = "" Then
GoTo ExitSub:
    
    End If
    ThisWorkbook.Sheets("inicio").Range("e12").Value = StrSavePath
ExitSub:
End Sub


Function BrowseForFolder(Optional iaa As Integer) As String
     
    Dim ShellApp As Object
     Dim OpenAt As Variant 'pasta inicial
     OpenAt = &H0 '
    'ssfDESKTOP           = 0x00,
    'ssfDESKTOPDIRECTORY  = 0x10,
    'ssfAPPDATA           = 0x1a,

     
    Set ShellApp = CreateObject("Shell.Application"). _
    BrowseForFolder(0, "Escolha a pasta", iaa, &H0)
     
    On Error Resume Next
    BrowseForFolder = ShellApp.self.Path
    On Error GoTo 0
     
    Select Case Mid(BrowseForFolder, 2, 1)
    Case Is = ":"
        If Left(BrowseForFolder, 1) = ":" Then
            BrowseForFolder = ""
        End If
    Case Is = "\"
        If Not Left(BrowseForFolder, 1) = "\" Then
            BrowseForFolder = ""
        End If
    Case Else
        BrowseForFolder = ""
    End Select
     
ExitFunction:
     
    Set ShellApp = Nothing
     
End Function
