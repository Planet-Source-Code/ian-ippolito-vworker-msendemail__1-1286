<div align="center">

## mSendEmail


</div>

### Description

If you have Outlook 98 you can send email using VB! Use this code for the basis of creating mailing programs!
 
### More Info
 
vcolEmailAddress--collection of string email address

vstrSubject--email subject

vstrBody--email body (use vbCrLf to create lf)

Requires outlook 98 installed on your machine.  Also, make sure you set a reference in your VB project to the Outlook 98 Type Library or this won't compile.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Ian Ippolito \(vWorker\)](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ian-ippolito-vworker.md)
**Level**          |Unknown
**User Rating**    |4.2 (67 globes from 16 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/ian-ippolito-vworker-msendemail__1-1286/archive/master.zip)





### Source Code

```
Sub mSendEmail(ByVal vcolEmailAddress As Collection, _
  ByVal vstrSubject As String, _
  ByVal vstrBody As String)
Dim ol As New Outlook.Application
Dim ns As Outlook.NameSpace
  'Return a reference to the MAPI layer
  Dim newMail As Outlook.MailItem
  'Create a new mail message item
  Set ns = ol.GetNamespace("MAPI")
  Set newMail = ol.CreateItem(olMailItem)
  'set properties
  With newMail
    'Add the subject of the mail message
    .Subject = vstrSubject
    'Create some body text
    .Body = vstrBody
    '**************
    'go through all
    'addresses passed in
    '**************
    Dim strEmailAddress As String
    Dim intIndex As Integer
    For intIndex = 1 To vcolEmailAddress.Count
      strEmailAddress = vcolEmailAddress.Item(intIndex)
      'Add a recipient and test to make sure that the
      'address is valid using the Resolve method
      With .Recipients.Add(strEmailAddress)
        .Type = olTo
        If Not .Resolve Then
          'MsgBox "Unable to resolve address.", vbInformation
          Debug.Print "Unable to resolve address " & strEmailAddress & "."
          'Exit Sub
        End If
      End With
    Next intIndex
'    'Attach a file as a link with an icon
'    With .Attachments.Add _
'      ("\\Training\training.xls", olByReference)
'      .DisplayName = "Training info"
'    End With
    'Send the mail message
    .Send
    End With
    'Release memory
    Set ol = Nothing
    Set ns = Nothing
    Set newMail = Nothing
End Sub
```

