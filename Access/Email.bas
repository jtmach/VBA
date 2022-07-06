Attribute VB_Name = "Email"
Option Compare Database
Option Explicit

#If VBA7 Then
  Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal lngMilliSeconds As Long)
#Else
  Private Declare Sub Sleep Lib "kernel32" (ByVal lngMilliSeconds As Long)
#End If

Public Function Test()
  Dim msgTo As String
  msgTo = "Test@gmail.com;"
  
  SendEmail msgTo, , "Test", "Test a break<br /> Link <a href='file://" & Replace(CurrentDb.Name, " ", "%20") & "'>File</a>", , True, False
End Function

Public Function GetOutlookApplication() As Outlook.Application
On Error Resume Next
  Dim outlookObj As Outlook.Application
  Set outlookObj = GetObject(, "Outlook.Application")
  
  If outlookObj Is Nothing Then ' Tries to get a currently running instance of Outlook, if one does not exist start a new one
    Set outlookObj = New Outlook.Application ' CreateObject("Outlook.Application")
  End If
  
  Set GetOutlookApplication = outlookObj
Exit Function
Err:
  MsgBox "Cannot access Outlook, please make sure Outlook is open before trying to import templates" & vbCrLf & Err.Description, vbOKOnly, "Error"
End Function

' Used for saving a sent item to a folder as a file
' Since we don't get any event when the message is sent, we need to look through the sent items to try to find a match
' This is not an optimal solution, but it works most of the time
Public Function SaveSentItem(msgSubject As String, folderPath As String)
On Error GoTo Err
  Dim oOutlook As Object
  Dim oOutlookNamespace As Object
  Set oOutlook = GetOutlookApplication()
  Set oOutlookNamespace = oOutlook.GetNamespace("MAPI")
  
  Dim foundItem As Object
  Dim foundItemIndex As Integer
  Dim oItems As Object
  Do While foundItem Is Nothing And foundItemIndex < 120 'Loop for 120 seconds if not found
    Set oItems = oOutlookNamespace.GetDefaultFolder(5).items '5 = olFolderSentMail
    oItems.sort "ReceivedTime", True
    If oItems(1).subject = msgSubject Then
      Set foundItem = oItems.Item(1)
    End If
    
    If foundItem Is Nothing Then
      Sleep 1000
    End If
    
    foundItemIndex = foundItemIndex + 1
    DoEvents
  Loop
  
  If Not foundItem Is Nothing Then
    Dim cleanedSubject As String
    cleanedSubject = Trim(SanitizeString(foundItem.subject))
    
    Dim fileName As String
    fileName = folderPath & Format(foundItem.ReceivedTime, "YYYY-MM-DD HHmmSS") & "-" & cleanedSubject & ".msg"
    foundItem.SaveAs fileName, 3 ' 3 = olMSG 'olHTML
        
    While Not foundItem.Saved
      DoEvents
    Wend
  Else
    MsgBox "Sent item not found, please save the email manually", vbOKOnly, "Sent item not found"
  End If
Exit Function
Err:
  Debug.Print Err.Description
  MsgBox "Unable to save or send e-mail" & vbCrLf & "Please save manually", vbOKOnly, "Sent item not found"
End Function

Public Function SendEmail(msgTo As String, Optional msgCc As String, Optional msgSubject As String, Optional msgBody As String, Optional AttachmentPath As String = "None", Optional ImportanceHigh As Boolean = False, Optional SendNow As Boolean = False)
  Dim objOutlook As Object
  Dim objOutlookMsg As Object
  Dim objOutlookRecip As Object
  Dim objOutlookAttach As Object
  Set objOutlook = GetOutlookApplication()
  Set objOutlookMsg = objOutlook.CreateItem(0) '0 = olMailItem
  
  With objOutlookMsg
    .To = msgTo
    .CC = msgCc
    .subject = msgSubject
    .BodyFormat = 2 '2 = olFormatHTML
    .body = msgBody
    .HTMLBody = msgBody
                    
    If AttachmentPath <> "None" Then
      'if your are not adding an attachment do not put anthing in this argument when calling the function
      .Attachments.Add AttachmentPath
    End If
    
    If ImportanceHigh = False Then
      .Importance = 1 '1 = olImportanceNormal
    Else
      .Importance = 2 '2 = olImportanceHigh
    End If
    
    'If ReplyToAddress <> "None" Then
      '.ReplyRecipients is read-only
      '.ReplyRecipients = ReplyToAddress
    'End If
    
    'Resolving the message removes all invalid or duplicated e-mail address
RestartResolve:
    Dim recipientIndex As Integer
    recipientIndex = 1
    Dim uniqueRecipientList As String
    For Each objOutlookRecip In .Recipients
      objOutlookRecip.Resolve
      If Not objOutlookRecip.Resolve Then
        objOutlookRecip.Delete
      End If
      
      'Only validate user entries
      If objOutlookRecip.addressentry.AddressEntryUserType = 0 Then ' 0 = olExchangeUserAddressEntry
        If InStr(uniqueRecipientList, "/" & objOutlookRecip.addressentry.getexchangeuser & "/") > 0 Then
          objOutlookRecip.Delete
          uniqueRecipientList = ""
          GoTo RestartResolve
        Else
          uniqueRecipientList = uniqueRecipientList & "/" & objOutlookRecip.addressentry.getexchangeuser & "/"
        End If
      End If
      recipientIndex = recipientIndex + 1
    Next

    If SendNow = False Then
      .display (False)
    ElseIf SendNow = True Then
      .send
    End If
  End With
  
  Set objOutlookMsg = Nothing
  Set objOutlook = Nothing
End Function

private Function SanitizeString(dirtyString As String) As String
  Dim cleanString As String
  cleanString = Replace(dirtyString, "\", "")
  cleanString = Replace(cleanString, ",", "")
  cleanString = Replace(cleanString, "/", "")
  cleanString = Replace(cleanString, "?", "")
  cleanString = Replace(cleanString, "*", "")
  cleanString = Replace(cleanString, ":", "")
  cleanString = Replace(cleanString, "<", "")
  cleanString = Replace(cleanString, ">", "")
  cleanString = Replace(cleanString, "|", "")
  cleanString = Replace(cleanString, """", "") 'Double quotes
    
  SanitizeString = cleanString
End Function