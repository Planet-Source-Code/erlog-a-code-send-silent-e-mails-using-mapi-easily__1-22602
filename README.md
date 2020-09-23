<div align="center">

## A\+ code: send silent e\-mails using mapi easily\!


</div>

### Description

This code will teach the beginner how to send anonymous and silent e-mails from within their VB apps. Great for silently mailing back a keylogging file! This is meant for the absolute beginner so don't bother pointng out that it is simple. If you use it there is no need to give me credit but please vote or say you like it!
 
### More Info
 
you need to have a mapi - enabled e-mail client (such as outlook express) and place the controls on a form (right click toolbar then microsoft mapi controls) with a command button.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[erLog](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/erlog.md)
**Level**          |Beginner
**User Rating**    |4.7 (109 globes from 23 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/erlog-a-code-send-silent-e-mails-using-mapi-easily__1-22602/archive/master.zip)





### Source Code

```
'***********************
'By littlegreenrussian *
'***********************
Private Sub Command1_Click() 'user clicks send
	On Error GoTo mailerr: 'go to the error handling bit if there is an error
		MAPISession1.SignOn 'sign on
If MAPISession1.SessionID <> 0 Then 'signed on
With MAPIMessages1
	.SessionID = MAPISession1.SessionID
	.Compose 'start a new message
.AttachmentName = "..." 'attachment name
	.AttachmentPathName = Text1 ' attachment path (get this from the text box or a default dirrectory)
	.RecipAddress = Text2 'set the receiver's email to the one they specified (again, text box or a default address)
.MsgSubject = "..." 'set the subject
.MsgNoteText = "............" 'message text
	.Send False 'don't display a dialog saying it was sent
		End With
			Exit Sub
				End If
	mailerr: 'error handling
		MsgBox "Error " & Err.Description
	End Sub
```

