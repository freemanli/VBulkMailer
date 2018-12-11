Sub SendMail()
	Set oMsg = New ExtMessage
	Call SetupMail()
	Dim hisLetter, findStr, xlskey, xlsRow, iniKey
	Dim i : i = 0
	Dim intCount : intCount = oExcel.rowData.Count
	Dim SendProgress : Set SendProgress = New IE_GUI
	With SendProgress
		.window("type") = IE_GUI_PROGRESS_BAR
		.dialog("title") = oLang("6001")
		.dialog("head") = oLang("6002")
		.dialog("body") = Replace(oLang("6003"), "%d",  intCount)
		.Show "SendProgress"
	End With
	' Bulk Senders
	If intCount > 10 Then
		oMsg.UpdateFields "urn:schemas:mailheader:List-Unsubscribe", "<mailto:" & oIni.parser("Sender")("ReplyTo") & ">"
		If UCase(oIni.parser("Letter")("Format")) = "HTML" Then
			sLetter = Replace(sLetter, "</body>", "<span style='font-size:small;float:right;'><a href='mailto:" & oIni.parser("Sender")("ReplyTo") & "?subject=unsubscribe'>" & oLang("6005") & "</a></span></body>")
		Else
			sLetter = sLetter & vbCrlf & oLang("6005") & " mailto:" & oIni.parser("Sender")("ReplyTo")
		End If
	End If
	With oMsg
		For Each xlsRow In oExcel.rowData
			i = i + 1
			If oExcel.rowData(xlsRow)(oIni.parser("ColumnName")("Email")) <> "" And SendProgress.window("exist") Then
				' Receiver
				.Receiver	.FormatEmail(oExcel.rowData(xlsRow)(oIni.parser("ColumnName")("Name")), oExcel.rowData(xlsRow)(oIni.parser("ColumnName")("Email"))), _
							"", _
							.FormatEmail(oIni.parser("Bcc")("Name"), oIni.parser("Bcc")("MailTo"))
				' Content
				hisLetter = sLetter
				For Each xlsKey In oExcel.rowData(xlsRow)
					findStr = oIni.parser("Letter")("PreTag") & xlsKey & oIni.parser("Letter")("PostTag")
					hisLetter = Replace(hisLetter, findStr, oExcel.rowData(xlsRow)(xlsKey))
				Next
				.Content 	oIni.parser("Letter")("Subject"), hisLetter,  oIni.parser("Letter")("Format")
				' Attachments
				.DeleteAllAttachments
				For Each iniKey In oIni.parser("Attachements")
					If Not .AddAttachment(oIni.parser("App")("MailFolder") & "\" & oIni.parser("Attachements")(iniKey)) Then
						WScript.Echo(iniKey & vbCrlf & oIni.parser("Attachements")(iniKey) & vbCrlf & oLang("5007"))
						Set SendProgress = Nothing
						Call EndProcess()
					End If
				Next
				For Each iniKey In oIni.parser("HisAttachements")
					If Not .AddAttachment(oIni.parser("App")("MailFolder") & "\" & _
									Replace(oIni.parser("HisAttachements")(iniKey), "%s", xlsRow)) Then
						WScript.Echo(iniKey & vbCrlf & oIni.parser("HisAttachements")(iniKey) & vbCrlf & oLang("6008"))
						Set SendProgress = Nothing
						Call EndProcess()
					End If
				Next
				' WScript.Echo oMsg.Msg.Attachments.Count
				With SendProgress
					.dialog("head") = oLang("6002") & Replace(oLang("6004"), "%d",  i) & " / " & Replace(oLang("6003"), "%d",  intCount) 
					.dialog("body") = oLang("5003") & """" & 	oExcel.rowData(xlsRow)(oIni.parser("ColumnName")("Name")) & """ &lt;"& _
															oExcel.rowData(xlsRow)(oIni.parser("ColumnName")("Email")) & "&gt;"
					.SetPct((i-1) * 100 / intCount)										
				End With
				.Send
				Randomize
				WScript.Sleep(Rnd * WAIT_TIME)
			End If
		Next
		SendProgress.SetPct(100)
		WScript.Sleep 2000
	End With
		
	Set SendProgress = Nothing
End Sub

Sub SendTestMail()
	Set oMsg = New ExtMessage
	Call SetupMail()
	Dim iniKey
	Dim SendProgress : Set SendProgress = New IE_GUI
	With SendProgress
		.window("type") = IE_GUI_PROGRESS_BAR
		.dialog("title") = oLang("5001")
		.dialog("head") = oLang("5002")
		.dialog("body") = oLang("5002") & oIni.parser("TestTo")("Name") & " " & oIni.parser("TestTo")("MailTo")
		.Show "SendProgress"
	End With
	With oMsg
		.Receiver	.FormatEmail(oIni.parser("TestTo")("Name"), oIni.parser("TestTo")("MailTo")), _
					"", _
					.FormatEmail(oIni.parser("Bcc")("Name"), oIni.parser("Bcc")("MailTo"))
		.Content 	oIni.parser("Letter")("Subject"), sLetter,  oIni.parser("Letter")("Format")
		.DeleteAllAttachments
		For Each iniKey In oIni.parser("Attachements")
			If Not .AddAttachment(oIni.parser("App")("MailFolder") & "\" & oIni.parser("Attachements")(iniKey)) Then
				WScript.Echo(iniKey & vbCrlf & oIni.parser("Attachements")(iniKey) & vbCrlf & oLang("5007"))
				Set SendProgress = Nothing
				Call EndProcess()
			End If
		Next
		SendProgress.SetPct(20)
		.Send
		SendProgress.SetPct(100)
		WScript.Sleep 200
	End With
	SendProgress.Close
	Set SendProgress =  Nothing
	Dim sShow : sShow = oLang("5003") & oIni.parser("TestTo")("Name") & " " & oIni.parser("TestTo")("MailTo") &vbCrlf & _
						oLang("5004") & vbCrlf & _
						oLang("5005") & vbCrlf & _
						fso.GetParentFolderName(WScript.ScriptFullName) & "\" & oIni.parser("App")("LogFile") & vbCrlf & _
						oLang("5006")
	If Popup(sShow, , oLang("5001"), vbYesNo + vbInformation)<> vbYes Then Call EndProcess()
End Sub

Sub SetupMail()
	With oMsg
		.Setup 	oIni.parser("SMTP"), _
				oIni.parser("BodyPart"), _
				oIni.parser("App")("LogFile") 
		.Transmitter 	.FormatEmail(oIni.parser("Sender")("Name"), oIni.parser("Sender")("From")) , _
						.FormatEmail("", oIni.parser("Sender")("ReplyTo"))
	End With
End Sub

