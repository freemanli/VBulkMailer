'################
'#### CLASSES ###
'################
' Class 	: ExtMessage
' Purpose	: To send emails ¤¤¤å
' Usage		: Set obj = New ExtMessage

' The MIT License (MIT)
' 
' Copyright (c) 2019 Freeman Li <freeman.tam@gmail.com>
' 
' Permission is hereby granted, free of charge, to any person obtaining a copy of
' this software and associated documentation files (the "Software"), to deal in
' the Software without restriction, including without limitation the rights to
' use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of
' the Software, and to permit persons to whom the Software is furnished to do so,
' subject to the following conditions:
' 
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
' 
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
' FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR
' COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
' IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN
' CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

' 	' ================================ Example 1 : The usage of ExtMessage ================================
'	Set smtp = CreateObject("Scripting.Dictionary")
'	smtp("sendusing") = 2
'	smtp("smtpserver") = "somesmtp.mail.com"
'	smtp("smtpserverport") = 25
'	smtp("smtpauthenticate") = 1
'	smtp("smtpusessl") = 1
'	smtp("smtpconnectiontimeout") = 30
'	smtp("sendusername") = "someone@somesmtp.com"
'	smtp("sendpassword") = "password"
'	Set bodyPart = CreateObject("Scripting.Dictionary")
'	bodyPart("ContentTransferEncoding") = "7bit"
'	bodyPart("Charset") = "big5"
'	
'	Dim eCDO
'	Set eCDO = New ExtMessage
'	
'	
'	eCDO.Setup smtp, bodyPart, "sender.log"
'	
'	With eCDO.Msg
'		.From 		= """TAM"" <tam.taipei@outlook.com>"
'		.To	  		= """You"" <someone@test.com>"
'		.Subject 	= "Test Email"
'		.TextBody 	="This is a test email."
'	End With
'	eCDO.AddAttachment "Popup.vb"			' Item(1)
'	eCDO.AddAttachment "33151643_3.jpg"		' Item(2)
'	eCDO.AddAttachment "test.pdf"			' Item(3)
'	'eCDO.Send
'	' Call WhatsVar(eCDO.Msg.Attachments.Item(1).Fields(2))
'	With eCDO.Msg.Attachments.Item(3)			' index only from 0 ~ 4
'												' 0 : filename 1 : attachment 2 : MIME type 3 attachment; filename=  4 MIME type; name=
'		WScript.Echo .Fields(0) & vbCrlf & _
'					.Fields(1) & vbCrlf & _
'					.Fields(2) & vbCrlf & _
'					.Fields(3) & vbCrlf & _
'					.Fields(4) & vbCrlf
'	End With
'	' WScript.Echo "Attachments Count : " & eCDO.Msg.Attachments.Count
'	' eCDO.Msg.Attachments.Item(3).Delete		' Not Support
'	eCDO.DeleteAllAttachments				' Delete All of the Attachments
'	' WhatsVar(eCDO.Msg.From)
'	eCDO.Send
'	
'	Set eCDO = Nothing
' ================================ End of Example 1 ================================

'################
'#### CLASSES ###
'################
Const CDO_CONFIG = "http://schemas.microsoft.com/cdo/configuration/"
Const EXTMESSAGE_TEXTBODY = "TEXT"
Const EXTMESSAGE_HTMLBODY = "HTML"

Class ExtMessage
	
	Private logFile, check
	
	Public Msg
	
	Public Sub Send()
		On Error Resume Next 
		Dim logMsg : logMsg = Now & vbTab
		Dim mailTo : mailTo = Msg.To
		' Send
		Msg.Send
		If Err.Number <> 0 Then
			logMsg = logMsg & Now & vbTab & "Failure" & vbTab & mailTo & vbTab
			logMsg = logMsg & "Error Number " & Err.Number & " : " & Replace(Err.Description, vbCrlf, "")
		Else
			logMsg = logMsg & Now & vbTab & "Success" & vbTab & mailTo
		End If
		logFile.WriteLine logMsg
	End Sub
	
	Public Function CheckEmailAddress(email)
		With New RegExp
			.Pattern    = "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,4}$"
			.Global     = False
			CheckEmailAddress = .Test(email)
		End With
	End Function
	
	Public Function FormatEmail(name, email)
		If name = "" Then
			FormatEmail = "<" & email & ">"
		Else
			FormatEmail = """" & name & """ <" & email & ">"
		End If
	End Function
	
	Public Sub Transmitter(sender, replyto)
		If sender<>"" Then Msg.From = sender
		If replyto<>"" Then Msg.ReplyTo = replyto
	End Sub
	
	Public Sub Receiver(recipient, cc, bcc)
		If recipient<>"" Then Msg.To = recipient
		If cc<>"" Then Msg.CC = cc
		If bcc<>"" Then Msg.BCC = bcc
	End Sub
	
	Public Sub Content(subject, body, format)
		With Msg
			If subject<>"" Then .Subject = subject
			If format = EXTMESSAGE_HTMLBODY Then 
				.HTMLBody = body
			Else
				.TextBody = body
			End If
		End With
	End Sub
	
	Public Function AddAttachment(relativePath)
		Dim absolutePath
		If Right(relativePath, 1) <> "\" Then
			With CreateObject("Scripting.FileSystemObject")
				absolutePath = .GetParentFolderName(WScript.ScriptFullName)& "\" & relativePath
				If .FileExists(absolutePath) Then
					Msg.AddAttachment absolutePath
					AddAttachment = True
				Else
					AddAttachment = False
				End If
			End With
		Else
			AddAttachment = True
		End If
	End Function
	
	Public Sub DeleteAllAttachments()
		If Msg.Attachments.Count > 0 Then Msg.Attachments.DeleteAll
	End Sub
	
	Public Sub UpdateFields(key, value)
		With Msg.Fields
			.Item(key) = value
			.Update
		End With
	End Sub
	
	Public Sub Setup(smtp, bodyPart, logPath)
		Dim key
		' Setup SMTP
		With Msg
			With .Configuration.Fields
				For Each key In smtp
					.Item(CDO_CONFIG & key) = smtp(key)
				Next
			End With
			.Configuration.Fields.Update
		End With
		' Setup BodyPart
		' WScript.Echo Msg.BodyPart.ContentTransferEncoding & vbCrlf & Msg.BodyPart.Charset
		If bodyPart.Exists("ContentTransferEncoding") Then 
			If bodyPart("ContentTransferEncoding") <> ""  Then Msg.BodyPart.ContentTransferEncoding = bodyPart("ContentTransferEncoding")
		End If
		If bodyPart.Exists("Charset") Then 
			If bodyPart("Charset")<> "" Then Msg.BodyPart.Charset = bodyPart("Charset")
		End If
		
		' Setup Log File
		With CreateObject("Scripting.FileSystemObject")
			Set logFile = .OpenTextFile(logPath, 8, True)
		End With
		'WhatsVar mailTo
	End Sub
	
	Private Sub Class_Initialize()
		Set Msg = CreateObject("CDO.Message")
	End Sub
	
	Private Sub Class_Terminate()
		On Error Resume Next 
		logFile.close
		Set logFile = Nothing
		Set Msg = Nothing
	End Sub

End Class


Sub WhatsVar(var)
	Dim strWhat, showValue 
	Dim VarTypeMeaning
	Select Case VarType(var)
		Case 0 		: showValue=False : VarTypeMeaning = "vbEmpty - Indicates Empty (uninitialized)"
		Case 1 		: showValue=False : VarTypeMeaning = "vbNull - Indicates Null (no valid data)"
		Case 2 		: showValue=True : VarTypeMeaning = "vbInteger - Indicates an integer"
		Case 3 		: showValue=True : VarTypeMeaning = "vbLong - Indicates a long integer"
		Case 4 		: showValue=True : VarTypeMeaning = "vbSingle - Indicates a single-precision floating-point number"
		Case 5 		: showValue=True : VarTypeMeaning = "vbDouble - Indicates a double-precision floating-point number"
		Case 6 		: showValue=True : VarTypeMeaning = "vbCurrency - Indicates a currency"
		Case 7 		: showValue=True : VarTypeMeaning = "vbDate - Indicates a date"
		Case 8 		: showValue=True : VarTypeMeaning = "vbString - Indicates a string"
		Case 9 		: showValue=False : VarTypeMeaning = "vbObject - Indicates an automation object"
		Case 10 	: showValue=True : VarTypeMeaning = "vbError - Indicates an error"
		Case 11 	: showValue=True : VarTypeMeaning = "vbBoolean - Indicates a boolean"
		Case 12 	: showValue=False : VarTypeMeaning = "vbVariant - Indicates a variant (used only with arrays of Variants)"
		Case 13 	: showValue=False : VarTypeMeaning = "vbDataObject - Indicates a data-access object"
		Case 17 	: showValue=True : VarTypeMeaning = "vbByte - Indicates a byte"
		Case 8192 	: showValue=False : VarTypeMeaning = "vbArray - Indicates an array"
		Case else 	: showValue=False : VarTypeMeaning = "unKnown"
	End Select
	strWhat = "TypeName = " & TypeName(var) & vbCrlf 
	strWhat = strWhat & "VarType = " & VarType(var) & vbCrlf
	strWhat = strWhat & "VarTypeMeaning = " & VarTypeMeaning & vbCrlf
	strWhat = strWhat & "IsObject = " & IsObject(var) & vbCrlf
	strWhat = strWhat & "IsNumeric = " & IsNumeric(var) & vbCrlf
	strWhat = strWhat & "IsNull = " & IsNull(var) & vbCrlf
	strWhat = strWhat & "IsEmpty = " & IsEmpty(var) & vbCrlf
	strWhat = strWhat & "IsArray = " & IsArray(var) & vbCrlf
	strWhat = strWhat & "IsDate = " & IsDate(var) & vbCrlf
	If showValue Then strWhat = strWhat & "Value = " & var & vbCrlf
	WScript.Echo strWhat
End Sub
