Option Explicit

'SHCreateDirectoryEx - Minimum operating systems: Windows 2000, Windows Millennium Edition
'Private Declare Function SHCreateDirectoryEx Lib "shell32" Alias "SHCreateDirectoryExA" (ByVal hwnd As Long, ByVal pszPath As String, ByVal psa As Any) As Long

#If VBA7 Then
    Private Declare PtrSafe Function SHCreateDirectoryEx Lib "shell32" Alias "SHCreateDirectoryExA" (ByVal hwnd As Long, ByVal pszPath As String, ByVal psa As Any) As Long
#Else
    Private Declare Function SHCreateDirectoryEx Lib "shell32" Alias "SHCreateDirectoryExA" (ByVal hwnd As Long, ByVal pszPath As String, ByVal psa As Any) As Long
#End If

Private Sub createfolders()
    Dim sPath As String
    Open "C:\config\folders.dat" For Input As #1
    Do Until EOF(1)
        Line Input #1, sPath
        sPath = Trim$(sPath)
        If Len(sPath) > 0 Then SHCreateDirectoryEx Me.hwnd, sPath, ByVal 0&
    Loop
    Close #1
End Sub

Private Sub tttt()
    Dim sssb As String
    sssb = Dir("C:\Users\Reinaldo\Desktop\teste2\3\2", vbDirectory)
    If Len(Dir("C:\Users\Reinaldo\Desktop\teste2\3\2", vbDirectory)) = 0 Then
        MkDir "C:\Users\Reinaldo\Desktop\teste2\3\2"
    End If

End Sub

Sub SaveAllEmails_ProcessAllSubFolders()
    'http://www.vbaexpress.com/kb/getarticle.php?kb_id=875
     
    Dim i               As Long
    Dim j               As Long
    Dim n               As Long
    Dim iaa             As Integer
    Dim StrSubject      As String
    Dim StrName         As String
    Dim StrFile         As String
    Dim StrReceived     As String
    Dim StrSavePath     As String
    Dim StrFolder       As String
    Dim StrFolderPath   As String
    Dim StrSaveFolder   As String
    Dim Prompt          As String
    Dim Title           As String
    Dim iNameSpace      As NameSpace
    Dim myOlApp         As Outlook.Application
    Dim SubFolder       As MAPIFolder
    Dim mItem           As MailItem
    Dim FSO             As Object
    Dim ChosenFolder    As Object
    Dim Folders         As New Collection
    Dim EntryID         As New Collection
    Dim StoreID         As New Collection
     
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set myOlApp = Outlook.Application
    Set iNameSpace = myOlApp.GetNamespace("MAPI")
    Set ChosenFolder = iNameSpace.PickFolder
    If ChosenFolder Is Nothing Then
        GoTo ExitSub:
    End If
     
    Prompt = "Entre com o caminho onde quer salvar"
    Title = "Pasta de Destino"
    StrSavePath = BrowseForFolder
    If StrSavePath = "" Then
        GoTo ExitSub:
    End If
    If Not Right(StrSavePath, 1) = "\" Then
        StrSavePath = StrSavePath & "\"
    End If
     
    Call GetFolder(Folders, EntryID, StoreID, ChosenFolder)
     
    iaa = 0
     
    For i = 1 To Folders.Count
        StrFolder = StripIllegalChar(Folders(i))
        n = InStr(3, StrFolder, "\") + 1
        StrFolder = Mid(StrFolder, n, 256)
        StrFolderPath = StrSavePath & StrFolder & "\"
        StrSaveFolder = Left(StrFolderPath, Len(StrFolderPath) - 1) & "\"
        
        'If Not FileFolderExists(StrFolderPath) Then
        'MkDir (StrFolderPath)
        'End If
        Call CreateSubDirectories(StrFolderPath)
        
        ' If Not FSO.FolderExists(StrFolderPath) Then
        '     FSO.CreateFolder (StrFolderPath)   '(Left(StrFolderPath, Len(StrFolderPath) - 1))
        '  End If
         
        Set SubFolder = myOlApp.Session.GetFolderFromID(EntryID(i), StoreID(i))
        On Error Resume Next
        For j = 1 To SubFolder.Items.Count
            Set mItem = SubFolder.Items(j)
            StrReceived = ArrangedDate(mItem.ReceivedTime)
            StrSubject = mItem.Subject
            StrName = StripIllegalChar(StrSubject)
            StrFile = StrSaveFolder & StrReceived & "" & StrName & ".msg"
            StrFile = Left(StrFile, 256)
            If Len(Dir(StrFile)) > 0 Then
                StrFile = Left(StrFile, Len(StrFile) - 4)
                StrFile = StrFile & " (" & iaa & ")" & ".msg" 'File exists
                iaa = iaa + 1
            End If
            mItem.SaveAs StrFile, 3
        Next j
        On Error GoTo 0
    Next i
     
ExitSub:
     
End Sub

Function StripIllegalChar(StrInput)
     
    Dim RegX            As Object
     
    Set RegX = CreateObject("vbscript.regexp")
     
    RegX.Pattern = "[\" & Chr(34) & "\!\@\#\$\%\^\&\*\(\)\=\+\|\[\]\{\}\`\'\;\:\<\>\?\/\,]"
    RegX.IgnoreCase = True
    RegX.Global = True
     
    StripIllegalChar = RegX.Replace(StrInput, "")
     
ExitFunction:
     
    Set RegX = Nothing
     
End Function

Function ArrangedDate(StrDateInput)
     
    Dim StrFullDate     As String
    Dim StrFullTime     As String
    Dim StrAMPM         As String
    Dim StrTime         As String
    Dim StrYear         As String
    Dim StrMonthDay     As String
    Dim StrMonth        As String
    Dim StrDay          As String
    Dim StrDate         As String
    Dim StrDateTime     As String
    Dim RegX            As Object
    
    Dim StrDataRei1     As String
    Dim StrDataRei2     As String
    Dim StrDiaRei       As String
    Dim StrMesRei       As String
    Dim StrAnoRei       As String
    Dim IntPosRei       As Integer
    
    
    StrDataRei1 = StrDateInput
    
    'StrDataRei1 = Replace(StrDataRei1, "#", "")                         ' 28/03/2013 16:05:31
    IntPosRei = InStr(StrDataRei1, " ")          ' 11
    StrDataRei2 = Left(StrDataRei1, IntPosRei - 1) ' 28/03/2013
    IntPosRei = InStr(StrDataRei2, "/")          ' 3
    StrDiaRei = Left(StrDataRei1, IntPosRei - 1) ' 28
    StrDataRei2 = Right(StrDataRei2, Len(StrDataRei2) - IntPosRei) ' 03/2013
    IntPosRei = InStr(StrDataRei2, "/")          ' 3
    StrMesRei = Left(StrDataRei2, IntPosRei - 1) ' 03
    StrAnoRei = Right(StrDataRei2, Len(StrDataRei2) - IntPosRei)
    'StrAnoRei = Replace(StrDataRei1, "#", "") ' 2013
    
     
    Set RegX = CreateObject("vbscript.regexp")
     
    If Not Left(StrDateInput, 2) = "10" And _
                                   Not Left(StrDateInput, 2) = "11" And _
                                   Not Left(StrDateInput, 2) = "12" Then
        StrDateInput = "0" & StrDateInput
    End If
     
    StrFullDate = Left(StrDateInput, 10)
     
    If Right(StrFullDate, 1) = " " Then
        StrFullDate = Left(StrDateInput, 9)
    End If
     
    StrFullTime = Replace(StrDateInput, StrFullDate & " ", "")
     
    If Len(StrFullTime) = 10 Then
        StrFullTime = "0" & StrFullTime
    End If
     
    StrAMPM = Right(StrFullTime, 2)
    StrTime = StrAMPM & "-" & Left(StrFullTime, 8)
    StrYear = Right(StrFullDate, 4)
    StrMonthDay = Replace(StrFullDate, "/" & StrYear, "")
    StrMonth = Left(StrMonthDay, 2)
    StrDay = Right(StrMonthDay, Len(StrMonthDay) - 3)
    If Len(StrDay) = 1 Then
        StrDay = "0" & StrDay
    End If
    StrDate = StrYear & "-" & StrMonth & "-" & StrDay ' & " - "
    StrDateTime = StrDate & "_" & StrTime
    RegX.Pattern = "[\:\/\ ]"
    RegX.IgnoreCase = True
    RegX.Global = True
     
    'ArrangedDate = RegX.Replace(StrDateTime, "-")
    ArrangedDate = StrAnoRei & "." & StrMesRei & "." & StrDiaRei & " - "
     
ExitFunction:
     
    Set RegX = Nothing
     
End Function

Sub GetFolder(Folders As Collection, EntryID As Collection, StoreID As Collection, Fld As MAPIFolder)
     
    Dim SubFolder       As MAPIFolder
     
    Folders.Add Fld.FolderPath
    EntryID.Add Fld.EntryID
    StoreID.Add Fld.StoreID
    For Each SubFolder In Fld.Folders
        GetFolder Folders, EntryID, StoreID, SubFolder
    Next SubFolder
     
ExitSub:
     
    Set SubFolder = Nothing
     
End Sub

Function BrowseForFolder(Optional OpenAt As String) As String
     
    Dim ShellApp As Object
     
    Set ShellApp = CreateObject("Shell.Application"). _
        BrowseForFolder(0, "Please choose a folder", 0, OpenAt)
     
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

Public Function FileFolderExists(strFullPath As String) As Boolean
    'Author       : Ken Puls (www.excelguru.ca)
    'Macro Purpose: Check if a file or folder exists
    On Error GoTo EarlyExit
    If Not Dir(strFullPath, vbDirectory) = vbNullString Then FileFolderExists = True
    
EarlyExit:
    On Error GoTo 0
End Function

Sub CreateSubDirectories(fullPath As String)
  
    Dim str As String
    Dim strArray As Variant
    Dim i As Long
    Dim basePath As String
    Dim newPath As String
  
    str = fullPath
  
    ' add trailing slash
    If Right$(str, 1) <> "\" Then
        str = str & "\"
    End If
  
    ' split string into array
    strArray = Split(str, "\")
  
    basePath = strArray(0) & "\"
  
    ' loop through array and create progressively
    ' lower level folders
    For i = 1 To UBound(strArray) - 1
        If Len(newPath) = 0 Then
            newPath = basePath & newPath & strArray(i) & "\"
        Else
            newPath = newPath & strArray(i) & "\"
        End If
  
        If Not FolderExists(newPath) Then
            MkDir newPath
        End If
    Next i
  
End Sub

Function FolderExists(ByVal strPath As String) As Boolean
    ' from http://allenbrowne.com
    On Error Resume Next
    FolderExists = ((GetAttr(strPath) And vbDirectory) = vbDirectory)
End Function


