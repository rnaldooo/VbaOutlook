Attribute VB_Name = "Módulo6"
Sub emailanexo()

 '030 Fetch email attributes
Dim ol As Outlook.Application
Dim msg As Outlook.MailItem
Dim myItems, myItem, myAttachments, myAttachment As Object
Dim myOrt As String
Set ol = New Outlook.Application
Set msg = ol.CreateItemFromTemplate(SSP & "\" & FN)
 
 '040
MySubject = msg.Subject
MySenderName = msg.SenderName
MyBody = msg.Body
MySentOn = msg.SentOn
 
 '050
myOrt = "C:\Temp\Scratch Pad\"
 'for all attachments do...
With ol
    For i = 1 To msg.Attachments.Count
         
         'save them to destination
        On Error GoTo 0
         'myAttachments(i).SaveAsFile myOrt & myAttachments(i).DisplayName '<---Type Mismatch
        myAttachments(i).SaveAsFile myOrt & "TEST99.doc" '<---Type Mismatch
         
         'add name and destination to message text 'I don't think I need this line (or next)
         'myItem.Body = myItem.Body & "File: " & myOrt & myAttachments(i).DisplayName & vbCrLf '
         
    Next i
End With



End Sub



Sub x(strSearchFolder As String, strOutputFolder As String)
     
    Dim i As Integer
    Dim i2 As Integer
    Dim strFile As String
    Dim objMsg As Object
     
    Dim ol As Outlook.Application
    Dim msg As Outlook.MailItem
     
    Set ol = New Outlook.Application
    
     strFile = Dir$(strSearchFolder & "\*.*")
   ' strFile = Dir$(strSearchFolder & "\*.MSG")
     
    Do While strFile <> vbNullString
    
    'objsmsg = strFile
    'If objMsg.Class = olMail Then
' Get the Attachments collection of the item.
'Set objAttachments = objMsg.Attachments
'lngCount = objAttachments.Count
'If lngCount > 0 Then
' We need to use a count down loop for
' removing items from a collection. Otherwise,
' the loop counter gets confused and only every
' other item is removed.
'For i = lngCount To 1 Step -1
' Save attachment before deleting from item.
' Get the file name.
'strFile = objAttachments.Item(i).FileName
         
        Set msg = ol.CreateItemFromTemplate(strSearchFolder & "\" & strFile)
         
         'for all attachments do...
        For i = 1 To msg.Attachments.Count
        ThisWorkbook.Sheets("anexos").Range("b" & i2 + 1 & "").Value = strSearchFolder & "\" & strFile
        ThisWorkbook.Sheets("anexos").Range("b" & i2 + 1 & "").Value = msg.Attachments.FileName
           ' Msg.Attachments(i).SaveAsFile strOutputFolder & Msg.Attachments(1).FileName
           i2 = i2 + 1
        Next i
         
        Set msg = Nothing
        strFile = Dir
         
    Loop
     
    Set ol = Nothing
     
End Sub


Sub ttttt1()

 Set OlApp = GetObject(, "Outlook.Application")
    If OlApp Is Nothing Then Err.Raise ERR_OUTLOOK_NOT_OPEN
    
    Set Eml = OlApp.CreateItemFromTemplate(MsgFilePath)
    For Each Attch In Eml.Attachments
       Attch.SaveAsFile SaveLocation & Attch.FileName
    Next
    
    Kill MsgFilePath



End Sub



Sub rrrein()

   Dim ia                As Integer
   Dim ilinha            As Integer
   Dim itotal            As Integer
   Dim socaminho         As String
   Dim sdcaminho         As String
   Dim snarquivo         As String
   Dim spcaminho         As String
   socaminho = ThisWorkbook.Sheets("inicio").Range("e3").Value
   'sdcaminho = ThisWorkbook.Sheets("inicio").Range("e12").Value
   'itotal = ThisWorkbook.Sheets("inicio").Range("e14").Value
   
   strSearchFolder = socaminho
  ' strOutputFolder = sdcaminho
   
   
   Dim i As Integer
    Dim i2 As Integer
    Dim strFile As String
    Dim objMsg As Object
     
    Dim ol As Outlook.Application
    'dim emll as
    Dim msg As Outlook.MailItem
     
    Set ol = New Outlook.Application
     
    strFile = Dir$(strSearchFolder & "\*.MSG")
    'Load mail message
   
    
     'strFile = Dir$(strSearchFolder & "\*.*")
    Do While strFile <> vbNullString
    
    'objsmsg = strFile
    'If objMsg.Class = olMail Then
' Get the Attachments collection of the item.
'Set objAttachments = objMsg.Attachments
'lngCount = objAttachments.Count
'If lngCount > 0 Then
' We need to use a count down loop for
' removing items from a collection. Otherwise,
' the loop counter gets confused and only every
' other item is removed.
'For i = lngCount To 1 Step -1
' Save attachment before deleting from item.
' Get the file name.
'strFile = objAttachments.Item(i).FileName
         
        Set msg = ol.CreateItemFromTemplate(strSearchFolder & "\" & strFile)
         
         'for all attachments do...
        For i = 1 To msg.Attachments.Count
        
        ThisWorkbook.Sheets("inicio").Range("e" & i2 + 15 & "").Value = strFile 'strSearchFolder & "\" & strFile
        ThisWorkbook.Sheets("inicio").Range("f" & i2 + 15 & "").Value = msg.Attachments(i).FileName
        ThisWorkbook.Sheets("inicio").Range("g" & i2 + 15 & "").Value = msg.SentOn
           ' Msg.Attachments(i).SaveAsFile strOutputFolder & Msg.Attachments(1).FileName
           i2 = i2 + 1
        Next i
         
        Set msg = Nothing
        strFile = Dir
         
    Loop
     
    Set ol = Nothing
   
   
   
   
 'Call x(socaminho, sdcaminho)
   
 '  For ia = 1 To itotal
  ' snarquivo = ThisWorkbook.Sheets("inicio").Range("f" & ia + 14 & "").Value
   
  '       Set OlApp = GetObject(, "Outlook.Application")
  '       If OlApp Is Nothing Then Err.Raise ERR_OUTLOOK_NOT_OPEN
 '        Set Eml = OlApp.Session.OpenSharedItem(socaminho & "\" & snarquivo)
  '      ' Set Eml = OlApp.CreateItemFromTemplate(snarquivo)
 '        For Each Attch In Eml.Attachments
 '        ThisWorkbook.Sheets("anexos").Range("a" & ilinha + 1 & "").Value = Attch.FileName
 '        ilinha = ilinha + 1
 '           'Attch.SaveAsFile SaveLocation & Attch.Filename
 '        Next
         
 '        Kill MsgFilePath
 '  ilinha = ilinha + 1
 '  Next ia
   'snarquivo = Replace(snarquivo, ".xls", "")
   'Dim ws As Worksheet, ss As Worksheet, FolderName As String, wb As Workbook
   ' Application.ScreenUpdating = False
    'FolderName = ThisWorkbook.Path
    
   ' Worksheets("Ident. Amostras").Copy
   ' Set wb = ActiveWorkbook

'Set OL = Nothing
End Sub

 
 Sub dfjslfl()
 
  
     Dim msg As System.Net.Mail.MailMessage =  New System.Net.Mail.MailMessage()
    msg.BodyEncoding = Encoding.UTF8
    msg.SubjectEncoding = Encoding.UTF8

    Dim plainView As AlternateView = AlternateView.CreateAlternateViewFromString(PrintablePage.StripHTML(body).Trim, System.Text.Encoding.UTF8, "text/plain")
    msg.AlternateViews.Add (plainView)
    Dim userState As Object = msg

    Response.Write (userState)
    
    
   Dim message As System.Net.Mail.MailMessage
   set System.Net.Mail.MailMessage.Load("test.eml", MessageFormat.Eml)
'Save to msg file
set message.Save("test.msg", MailMessageSaveType.OutlookMessageFormat)

 
 End Sub
 
 
 
Sub StripAttachments()
 
Dim objOL As Outlook.Application
Dim objMsg As Object
Dim objAttachments As Outlook.Attachments
Dim objSelection As Outlook.Selection
Dim i As Long
Dim lngCount As Long
Dim strFile As String
Dim strFolder As String
Dim p As Integer
 
On Error Resume Next
 
' Instantiate an Outlook Application object.
Set objOL = CreateObject("Outlook.Application")
' Get the collection of selected objects.
Set objSelection = objOL.ActiveExplorer.Selection
 
' Get the Temp folder.
' strFolder = GetTempDir()
' If strFolder = "" Then
' MsgBox "Could not get Temp folder", vbOKOnly
' GoTo ExitSub
' End If
 
' Check each selected item for attachments.
' If attachments exist, save them to the Temp
' folder and strip them from the item.
For Each objMsg In objSelection
' This code only strips attachments from mail items.
If objMsg.Class = olMail Then
' Get the Attachments collection of the item.
Set objAttachments = objMsg.Attachments
lngCount = objAttachments.Count
If lngCount > 0 Then
' We need to use a count down loop for
' removing items from a collection. Otherwise,
' the loop counter gets confused and only every
' other item is removed.
For i = lngCount To 1 Step -1
' Save attachment before deleting from item.
' Get the file name.
strFile = objAttachments.Item(i).FileName
' Combine with the path to the Temp folder.
' strFile = strFolder & strFile
' Save the attachment as a file.
objAttachments.Item(i).SaveAsFile "C:\Email Attachments\" & strFile
' Delete the attachment.
objAttachments.Item(i).Delete
p = p + 1
Next i
End If
objMsg.Save
End If
Next objMsg
 
If p > 0 Then
varResponse = MsgBox("I found " & p & " attached files." _
& vbCrLf & "I have saved them into the C:\Email Attachments folder." _
& vbCrLf & vbCrLf & "Would you like to view the files now?" _
, vbQuestion + vbYesNo, "Finished!")
' Open Windows Explorer to display saved files if user chooses
If varResponse = vbYes Then
Shell "Explorer.exe /e,C:\Email Attachments", vbNormalFocus
End If
Else
MsgBox "I didn't find any attached files in your mail.", vbInformation, "Finished!"
End If
 
 
 
ExitSub:
Set objAttachments = Nothing
Set objMsg = Nothing
Set objSelection = Nothing
Set objOL = Nothing
End Sub


Sub SaveLAttachmentsToFoldfsdfer()
 ' This Outlook macro checks a named subfolder in the Outlook Inbox
 ' (here the "Sales Reports" folder) for messages with attached
 ' files of a specific type (here file with an "xls" extension)
 ' and saves them to disk. Saved files are timestamped. The user
 ' can choose to view the saved files in Windows Explorer.
 ' NOTE: make sure the specified subfolder and save folder exist
 ' before running the macro.
 On Error GoTo SaveLAttachmentsToFolder_err
 ' Declare variables
 Dim ns As Namespace
 Dim Inbox As MAPIFolder
 Dim SubFolder As MAPIFolder
 Dim Item As Object
 Dim Atmt As Attachment
 Dim FileName As String
 Dim i As Integer
 Dim varResponse As VbMsgBoxResult
 Set ns = GetNamespace("MAPI")
 Set Inbox = ns.GetDefaultFolder(olFolderInbox)
 Set SubFolder = Inbox.Folders("Prime") ' Enter correct subfolder name.
 i = 0
 ' Check subfolder for messages and exit of none found
 If SubFolder.Items.Count = 0 Then
 MsgBox "There are no messages in the Sales Reports folder.", vbInformation, _
 "Nothing Found"
 Exit Sub
 End If
 ' Check each message for attachments
 For Each Item In SubFolder.Items
 For Each Atmt In Item.Attachments
 ' Check filename of each attachment and save if it has "xls" extension
 If Right(Atmt.FileName, 3) = "PDF" Then
 ' This path must exist! Change folder name as necessary.
 FileName = "g:\1 EXPENSES\DLLFeebillings\L\Price history\" & _
 Format(Item.CreationTime, "yyyymmdd_hhnnss_") & Atmt.FileName
 Atmt.SaveAsFile FileName
 i = i + 1
 End If
 Next Atmt
 Next Item
 ' Show summary message
 If i > 0 Then
 varResponse = MsgBox("I found " & i & " attached files." _
 & vbCrLf & "I have saved them into the g:\1 EXPENSES\DLLFeebillings\L\Price history" _
 & vbCrLf & vbCrLf & "Would you like to view the files now?" _
 , vbQuestion + vbYesNo, "Finished!")
 ' Open Windows Explorer to display saved files if user chooses
 If varResponse = vbYes Then
 Shell "Explorer.exe /e,g:\1 EXPENSES\DLLFeebillings\L\Price history\", vbNormalFocus
 End If
 Else
 MsgBox "I didn't find any attached files in your mail.", vbInformation, "Finished!"
 End If
 ' Clear memory
SaveLAttachmentsToFolder_exit:
 Set Atmt = Nothing
 Set Item = Nothing
 Set ns = Nothing
 Exit Sub
 ' Handle Errors
SaveLAttachmentsToFolder_err:
 MsgBox "An unexpected error has occurred." _
 & vbCrLf & "Please note and report the following information." _
 & vbCrLf & "Macro Name: GetAttachments" _
 & vbCrLf & "Error Number: " & Err.Number _
 & vbCrLf & "Error Description: " & Err.Description _
 , vbCritical, "Error!"
 Resume SaveLAttachmentsToFolder_exit
 End Sub
 

Option Explicit


Sub GetAttachmfddfents()
' This Outlook macro checks a the Outlook Inbox for messages
' with attached files (of any type) and saves them to disk.
' NOTE: make sure the specified save folder exists before
' running the macro.
    On Error GoTo GetAttachments_err
' Declare variables
    Dim ns As Namespace
    Dim Inbox As MAPIFolder
    Dim Item As Object
    Dim Atmt As Attachment
    Dim FileName As String
    Dim i As Integer
    Set ns = GetNamespace("MAPI")
    Set Inbox = ns.GetDefaultFolder(olFolderInbox)
    i = 0
' Check Inbox for messages and exit of none found
    If Inbox.Items.Count = 0 Then
        MsgBox "There are no messages in the Inbox.", vbInformation, _
               "Nothing Found"
        Exit Sub
    End If
' Check each message for attachments
    For Each Item In Inbox.Items
' Save any attachments found
        For Each Atmt In Item.Attachments
        ' This path must exist! Change folder name as necessary.
            FileName = "C:\Email Attachments\" & Atmt.FileName
            Atmt.SaveAsFile FileName
            i = i + 1
         Next Atmt
    Next Item
' Show summary message
    If i > 0 Then
        MsgBox "I found " & i & " attached files." _
        & vbCrLf & "I have saved them into the C:\Email Attachments folder." _
        & vbCrLf & vbCrLf & "Have a nice day.", vbInformation, "Finished!"
    Else
        MsgBox "I didn't find any attached files in your mail.", vbInformation, "Finished!"
    End If
' Clear memory
GetAttachments_exit:
    Set Atmt = Nothing
    Set Item = Nothing
    Set ns = Nothing
    Exit Sub
' Handle errors
GetAttachments_err:
    MsgBox "An unexpected error has occurred." _
        & vbCrLf & "Please note and report the following information." _
        & vbCrLf & "Macro Name: GetAttachments" _
        & vbCrLf & "Error Number: " & Err.Number _
        & vbCrLf & "Error Description: " & Err.Description _
        , vbCritical, "Error!"
    Resume GetAttachments_exit
End Sub

Sub SaveAttachmentsToFadfolder()
' This Outlook macro checks a named subfolder in the Outlook Inbox
' (here the "Sales Reports" folder) for messages with attached
' files of a specific type (here file with an "xls" extension)
' and saves them to disk. Saved files are timestamped. The user
' can choose to view the saved files in Windows Explorer.
' NOTE: make sure the specified subfolder and save folder exist
' before running the macro.
    On Error GoTo SaveAttachmentsToFolder_err
' Declare variables
    Dim ns As Namespace
    Dim Inbox As MAPIFolder
    Dim SubFolder As MAPIFolder
    Dim Item As Object
    Dim Atmt As Attachment
    Dim FileName As String
    Dim i As Integer
    Dim varResponse As VbMsgBoxResult
    Set ns = GetNamespace("MAPI")
    Set Inbox = ns.GetDefaultFolder(olFolderInbox)
    Set SubFolder = Inbox.Folders("Sales Reports") ' Enter correct subfolder name.
    i = 0
' Check subfolder for messages and exit of none found
    If SubFolder.Items.Count = 0 Then
        MsgBox "There are no messages in the Sales Reports folder.", vbInformation, _
               "Nothing Found"
        Exit Sub
    End If
' Check each message for attachments
    For Each Item In SubFolder.Items
        For Each Atmt In Item.Attachments
' Check filename of each attachment and save if it has "xls" extension
            If Right(Atmt.FileName, 3) = "xls" Then
            ' This path must exist! Change folder name as necessary.
                FileName = "C:\Email Attachments\" & _
                    Format(Item.CreationTime, "yyyymmdd_hhnnss_") & Atmt.FileName
                Atmt.SaveAsFile FileName
                i = i + 1
            End If
        Next Atmt
    Next Item
' Show summary message
    If i > 0 Then
        varResponse = MsgBox("I found " & i & " attached files." _
        & vbCrLf & "I have saved them into the C:\Email Attachments folder." _
        & vbCrLf & vbCrLf & "Would you like to view the files now?" _
        , vbQuestion + vbYesNo, "Finished!")
' Open Windows Explorer to display saved files if user chooses
        If varResponse = vbYes Then
            Shell "Explorer.exe /e,C:\Email Attachments", vbNormalFocus
        End If
    Else
        MsgBox "I didn't find any attached files in your mail.", vbInformation, "Finished!"
    End If
' Clear memory
SaveAttachmentsToFolder_exit:
    Set Atmt = Nothing
    Set Item = Nothing
    Set ns = Nothing
    Exit Sub
' Handle Errors
SaveAttachmentsToFolder_err:
    MsgBox "An unexpected error has occurred." _
        & vbCrLf & "Please note and report the following information." _
        & vbCrLf & "Macro Name: GetAttachments" _
        & vbCrLf & "Error Number: " & Err.Number _
        & vbCrLf & "Error Description: " & Err.Description _
        , vbCritical, "Error!"
    Resume SaveAttachmentsToFolder_exit
End Sub

Option Explicit


Sub GetAttachments()
' This Outlook macro checks a the Outlook Inbox for messages
' with attached files (of any type) and saves them to disk.
' NOTE: make sure the specified save folder exists before
' running the macro.
    On Error GoTo GetAttachments_err
' Declare variables
    Dim ns As Namespace
    Dim Inbox As MAPIFolder
    Dim Item As Object
    Dim Atmt As Attachment
    Dim FileName As String
    Dim i As Integer
    Set ns = GetNamespace("MAPI")
    Set Inbox = ns.GetDefaultFolder(olFolderInbox)
    i = 0
' Check Inbox for messages and exit of none found
    If Inbox.Items.Count = 0 Then
        MsgBox "There are no messages in the Inbox.", vbInformation, _
               "Nothing Found"
        Exit Sub
    End If
' Check each message for attachments
    For Each Item In Inbox.Items
' Save any attachments found
        For Each Atmt In Item.Attachments
        ' This path must exist! Change folder name as necessary.
            FileName = "C:\Email Attachments\" & Atmt.FileName
            Atmt.SaveAsFile FileName
            i = i + 1
         Next Atmt
    Next Item
' Show summary message
    If i > 0 Then
        MsgBox "I found " & i & " attached files." _
        & vbCrLf & "I have saved them into the C:\Email Attachments folder." _
        & vbCrLf & vbCrLf & "Have a nice day.", vbInformation, "Finished!"
    Else
        MsgBox "I didn't find any attached files in your mail.", vbInformation, "Finished!"
    End If
' Clear memory
GetAttachments_exit:
    Set Atmt = Nothing
    Set Item = Nothing
    Set ns = Nothing
    Exit Sub
' Handle errors
GetAttachments_err:
    MsgBox "An unexpected error has occurred." _
        & vbCrLf & "Please note and report the following information." _
        & vbCrLf & "Macro Name: GetAttachments" _
        & vbCrLf & "Error Number: " & Err.Number _
        & vbCrLf & "Error Description: " & Err.Description _
        , vbCritical, "Error!"
    Resume GetAttachments_exit
End Sub

Sub SaveAttachmenaaatsToFolder()
' This Outlook macro checks a named subfolder in the Outlook Inbox
' (here the "Sales Reports" folder) for messages with attached
' files of a specific type (here file with an "xls" extension)
' and saves them to disk. Saved files are timestamped. The user
' can choose to view the saved files in Windows Explorer.
' NOTE: make sure the specified subfolder and save folder exist
' before running the macro.
    On Error GoTo SaveAttachmentsToFolder_err
' Declare variables
    Dim ns As Namespace
    Dim Inbox As MAPIFolder
    Dim SubFolder As MAPIFolder
    Dim Item As Object
    Dim Atmt As Attachment
    Dim FileName As String
    Dim i As Integer
    Dim i2 As Integer
     Dim i3 As Integer
    Dim varResponse As VbMsgBoxResult
    Set ns = GetNamespace("MAPI")
    Set Inbox = ns.GetDefaultFolder(olFolderInbox)
    Set SubFolder = Inbox.Folders("tessstee") ' Enter correct subfolder name.
    i = 0
' Check subfolder for messages and exit of none found
    If SubFolder.Items.Count = 0 Then
        MsgBox "There are no messages in the Sales Reports folder.", vbInformation, _
               "Nothing Found"
        Exit Sub
    End If
' Check each message for attachments
Dim iaa As Integer
iaa = SubFolder.Items.Count
    For Each Item In SubFolder.Items
    i3 = 1
        For Each Atmt In Item.Attachments
        
        ThisWorkbook.Sheets("inicio").Range("e" & i2 + 15 & "").Value = Item.ConversationTopic 'strSearchFolder & "\" & strFile
        ThisWorkbook.Sheets("inicio").Range("f" & i2 + 15 & "").Value = Item.Attachments(i3).FileName
        ThisWorkbook.Sheets("inicio").Range("g" & i2 + 15 & "").Value = Atmt.FileName
        i3 = i3 + 1
        i2 = i2 + 1
        
' Check filename of each attachment and save if it has "xls" extension
        '    If Right(Atmt.FileName, 3) = "xls" Then
            ' This path must exist! Change folder name as necessary.
         '       FileName = "C:\Email Attachments\" & _
                    Format(Item.CreationTime, "yyyymmdd_hhnnss_") & Atmt.FileName
         '       Atmt.SaveAsFile FileName
         '       i = i + 1
          '  End If
        Next Atmt
    Next Item
' Show summary message
    If i > 0 Then
        varResponse = MsgBox("I found " & i & " attached files." _
        & vbCrLf & "I have saved them into the C:\Email Attachments folder." _
        & vbCrLf & vbCrLf & "Would you like to view the files now?" _
        , vbQuestion + vbYesNo, "Finished!")
' Open Windows Explorer to display saved files if user chooses
        If varResponse = vbYes Then
            Shell "Explorer.exe /e,C:\Email Attachments", vbNormalFocus
        End If
    Else
        MsgBox "I didn't find any attached files in your mail.", vbInformation, "Finished!"
    End If
' Clear memory
SaveAttachmentsToFolder_exit:
    Set Atmt = Nothing
    Set Item = Nothing
    Set ns = Nothing
    Exit Sub
' Handle Errors
SaveAttachmentsToFolder_err:
    MsgBox "An unexpected error has occurred." _
        & vbCrLf & "Please note and report the following information." _
        & vbCrLf & "Macro Name: GetAttachments" _
        & vbCrLf & "Error Number: " & Err.Number _
        & vbCrLf & "Error Description: " & Err.Description _
        , vbCritical, "Error!"
    Resume SaveAttachmentsToFolder_exit
End Sub



Sub exportemail1()
'Code is free to use
'Do not remove comments
'for any further help, please visit http://findsarfaraz.blogspot.com
'‘for email me findsarfaraz@gmail.com
On Error Resume Next
Dim emailcount As Integer
Dim OLF As Outlook.MAPIFolder
Dim ol As New Outlook.Application
'OLF is declared as mapi folder to decide'which folder you want to target
Set OLF = ol.Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox)
'OLF.Items.Count provide the number of mail present in inbox
emailcount = OLF.Items.Count
'by i=1 we 'are initializing the variable by value 1'else it will throw an error
i = 1
'Here I am using do while - loop to'Browse through all mail
Do While i <= emailcount
Sheet3.Cells(i + 1, 1) = OLF.Items(i).SenderEmailAddress
Sheet3.Cells(i + 1, 2) = OLF.Items(i).SenderName
Sheet3.Cells(i + 1, 3) = OLF.Items(i).SentOn
i = i + 1
Loop
'Like SenderEmailAddress SenderName SentOn'we can use other properties of email to get other 'details
Set OLF = Nothing
Set ol = Nothing
End Sub





Sub rrreadsadin()

   Dim ia                As Integer
   Dim ilinha            As Integer
   Dim itotal            As Integer
   Dim socaminho         As String
   Dim sdcaminho         As String
   Dim snarquivo         As String
   Dim spcaminho         As String
   socaminho = ThisWorkbook.Sheets("inicio").Range("e3").Value
   'sdcaminho = ThisWorkbook.Sheets("inicio").Range("e12").Value
   'itotal = ThisWorkbook.Sheets("inicio").Range("e14").Value
   
   strSearchFolder = socaminho
  ' strOutputFolder = sdcaminho
   
   
   Dim i As Integer
    Dim i2 As Integer
    Dim strFile As String
    Dim objMsg As Object
     
    Dim ol As Outlook.Application
    'dim emll as
    Dim msg As Outlook.MailItem
     
    Set ol = New Outlook.Application
     
    strFile = Dir$(strSearchFolder & "\*.eml")
    'Load mail message
   
    
     'strFile = Dir$(strSearchFolder & "\*.*")
    Do While strFile <> vbNullString
    

         
        Set msg = ol.CreateItemFromTemplate(strSearchFolder & "\" & strFile)
         
         'for all attachments do...
        For i = 1 To msg.Attachments.Count
        
        ThisWorkbook.Sheets("inicio").Range("e" & i2 + 15 & "").Value = strFile 'strSearchFolder & "\" & strFile
        ThisWorkbook.Sheets("inicio").Range("f" & i2 + 15 & "").Value = msg.Attachments(i).FileName
        ThisWorkbook.Sheets("inicio").Range("g" & i2 + 15 & "").Value = msg.SentOn
           ' Msg.Attachments(i).SaveAsFile strOutputFolder & Msg.Attachments(1).FileName
           i2 = i2 + 1
        Next i
         
        Set msg = Nothing
        strFile = Dir
         
    Loop
     
    Set ol = Nothing
   
   
   
   
 'Call x(socaminho, sdcaminho)
   
 '  For ia = 1 To itotal
  ' snarquivo = ThisWorkbook.Sheets("inicio").Range("f" & ia + 14 & "").Value
   
  '       Set OlApp = GetObject(, "Outlook.Application")
  '       If OlApp Is Nothing Then Err.Raise ERR_OUTLOOK_NOT_OPEN
 '        Set Eml = OlApp.Session.OpenSharedItem(socaminho & "\" & snarquivo)
  '      ' Set Eml = OlApp.CreateItemFromTemplate(snarquivo)
 '        For Each Attch In Eml.Attachments
 '        ThisWorkbook.Sheets("anexos").Range("a" & ilinha + 1 & "").Value = Attch.FileName
 '        ilinha = ilinha + 1
 '           'Attch.SaveAsFile SaveLocation & Attch.Filename
 '        Next
         
 '        Kill MsgFilePath
 '  ilinha = ilinha + 1
 '  Next ia
   'snarquivo = Replace(snarquivo, ".xls", "")
   'Dim ws As Worksheet, ss As Worksheet, FolderName As String, wb As Workbook
   ' Application.ScreenUpdating = False
    'FolderName = ThisWorkbook.Path
    
   ' Worksheets("Ident. Amostras").Copy
   ' Set wb = ActiveWorkbook

'Set OL = Nothing
End Sub

 


