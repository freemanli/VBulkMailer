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
Option Explicit
Include("include\Popup.vb")
Include("include\IniFile.vb")
Include("include\IE_GUI.vb")
Include("include\xlsreader.vb")
Include("include\ExtMessage.vb")
Include("include\EditIni.vb")
Const WAIT_TIME = 5000
' Declare Global Variable
Dim oIni, oLang, oExcel, oMsg, sLetter, fso, procGUI

Call Main

Sub Main()
	Dim execStep, sBody, i
	Dim mailerTitle : mailerTitle = "Mailer"
	
	If Not Initialize() Then Call EditIni()
	
	Set procGUI = New IE_GUI
	With procGUI
		.window("type") = IE_GUI_HTML
		.dialog("title") = mailerTitle
		.dialog("Width") = 200
		.dialog("Height") = 480
		.dialog("head") = 	"<style>button{width:180px;height:60px;margin:5px 15px;}</style>" & _
							"<script>function myFunction(a) {document.title='" & mailerTitle & "_' + a;}</script>"
		sBody = ""
		For i = 0 To 5
			sBody = sBody & "<button id='" & "step_" & i & "' "
			sBody = sBody & "class='step' " 
			sBody = sBody & "onclick='myFunction(" & i & ")'>"
			sBody = sBody & Replace(olang( (i+2) & "000"), "  ", "<br/>")
			sBody = sBody &	"</button>"
		Next
 		.dialog("body") = 	sBody 
 		.Show "procGUI"
		.Scroll = False
		While Not IsEmpty(.Title)
			'WhatsVar .Title
			If .Title <> mailerTitle Then
				execStep = Replace(.Title, mailerTitle & "_", "")
				.Visible = False
				.Title = mailerTitle
				Select Case execStep
					case "0"
						Call EditIni()
					case "1"
						Call ShowExcelData()
					case "2"
						Call ShowLetter()
					case "3"
						Call SendTestMail()
					case "4"
						Call SendMail()
						.Close
					case else
						.Close
				End Select
				.Visible = True
				.Activate "iexplore.exe"
			End If
			WScript.Sleep 200
		WEnd
	End With
	WScript.Sleep 200
	Call EndProcess()
End Sub

Sub EndProcess()
	On Error Resume Next
	With fso
		Dim destinationPath : destinationPath = .GetSpecialFolder(2) & "\"
		.DeleteFile destinationPath & "mail_error.png", True
		.DeleteFile destinationPath & "mail_okay.png", True
		.DeleteFile destinationPath & "mail_progress.gif", True
	End With
	procGUI.close
	Set procGUI = Nothing
	Set oExcel = Nothing
	Set oIni = Nothing
	Set oLang = Nothing
	Set oMsg = Nothing
	Set fso = Nothing
	WScript.Quit
End Sub

Sub ShowIni()
	Dim iniSec, iniKey
	Dim iniStr : iniStr = ""
	For Each iniSec In oIni.parser
		iniStr = iniStr & "[" & iniSec & "]" & vbLF & vbCR
		For Each iniKey In oIni.parser(iniSec)
			iniStr = iniStr & vbTab & iniKey & " = " & oIni.parser(iniSec)(iniKey) & vbLF & vbCR
		Next
	Next
	iniStr = iniStr & vbCrlf & oLang("2002")
	If Popup(iniStr, , oLang("2001"), 4+32) <> vbYes Then Call EndProcess()
End Sub

Sub ShowExcelData
	Dim xlsRow, xlsKey, xlsStr
	Dim i : i = oIni.parser("MailTo")("StartNo")
	If oIni.parser("MailTo")("StartNo") = "" Then i = 1
	For xlsRow = i To i+2 
'		WScript.Echo oIni.parser("MailTo")("StartNo")
		If oExcel.rowData.Exists(xlsRow) Then
			xlsStr = Replace(oLang("3002"), "%d", xlsRow) & vbCrlf
			For Each xlsKey In oExcel.rowData(xlsRow)
				xlsStr = xlsStr & vbTab & xlsKey & " = " & oExcel.rowData(xlsRow)(xlsKey) & vbCrlf
			Next
			xlsStr = xlsStr & vbCrlf & Replace(oLang("3003"), "%d", xlsRow)
			If Popup(xlsStr, , oLang("3001"), 4+32) <> vbYes Then Call EndProcess()
		End If
	Next
End Sub

Sub ShowLetter()
	Dim xlsRow, xlsKey
	Dim hisLetter, strFind
	Dim oLetter : Set oLetter = New IE_GUI
	Dim i : i = oIni.parser("MailTo")("StartNo")
	If oIni.parser("MailTo")("StartNo") = "" Then i = 1
	For xlsRow = i To i+2 
		If oExcel.rowData.Exists(xlsRow) Then
			hisLetter = sLetter
			For Each xlsKey In oExcel.rowData(xlsRow)
				strFind = oIni.parser("Letter")("PreTag") & xlsKey & oIni.parser("Letter")("PostTag")
				hisLetter = Replace(hisLetter, strFind, oExcel.rowData(xlsRow)(xlsKey))
			Next
			If oIni.parser("Letter")("Format") = "TEXT" Then hisLetter = "<pre>" & hisLetter & "</pre>"
			With oLetter
				.window("type") = IE_GUI_HTML
				.dialog("title") = Replace(oLang("4002"), "%d", xlsRow)
				.dialog("Top") = 0
				.dialog("Left") = 0
				.dialog("Width") = 960
				.dialog("Height") = 1040
				.dialog("head") = ""
				.dialog("body") = hisLetter
				.Show "oLetter"
			End With
			'statusBar.UpdateRegion 1, myLetter
			If Popup(Replace(oLang("4003"), "%d", xlsRow), , oLang("4001"), 4+32) <> vbYes Then 
				oLetter.Close
				Set oLetter = Nothing
				Call EndProcess()
			End If
			oLetter.Close
		End If
	Next
	Set oLetter = Nothing
End Sub

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

Function Initialize()
	Dim initialProgess : Set initialProgess = New IE_GUI
	Set fso = CreateObject("Scripting.FileSystemObject")
	' copy image files
	With fso
		Dim sourcePath : sourcePath = .GetParentFolderName(WScript.ScriptFullName)& "\"
		Dim destinationPath : destinationPath = .GetSpecialFolder(2) & "\"
		.CopyFile sourcePath & "mail_error.png", destinationPath & "mail_error.png"
		.CopyFile sourcePath & "mail_okay.png", destinationPath & "mail_okay.png"
		.CopyFile sourcePath & "mail_progress.gif", destinationPath & "mail_progress.gif"
	End With
	Call ParseIni()
	Call ParseLang()
	With initialProgess
		.window("type") 	= IE_GUI_PROGRESS_BAR
		.dialog("title") 	= oLang("1001")
		.dialog("head") 	= oLang("1002")
		.dialog("body") 	= oLang("1004")
		.Show "initialProgess"
	End With
	initialProgess.SetPct(10)
	WScript.Sleep 100
	Initialize = False
	If ReadLetter() Then 
		initialProgess.SetPct(30)
		WScript.Sleep 100
		initialProgess.dialog("body") = oLang("1003")
		If ParseExcel(initialProgess, 30) Then Initialize = True
	End If
	initialProgess.Close
	Set initialProgess = Nothing
End Function

Function ReadLetter()
	Dim letterPath : letterPath = oIni.parser("App")("MailFolder") & "\" & oIni.parser("Letter")("File")
	If FileExists(letterPath) Then
		sLetter = CreateObject("Scripting.FileSystemObject").OpenTextFile(letterPath, 1, False).ReadAll
		ReadLetter = True
	Else
		ReadLetter = False
	End If
End Function

Function ParseExcel(prog, pct)
	Set oExcel = New ExcelFile
	oExcel.SetProgress prog, pct
	ParseExcel = oExcel.Read(oIni.parser("App")("MailFolder") & "\" & oIni.parser("MailTo")("File"), _
				oIni.parser("MailTo")("Worksheet"), _
				oIni.parser("MailTo")("StartNo"), _
				oIni.parser("MailTo")("EndNo"))
End Function

Sub ParseLang()
	Set oLang = CreateObject("Scripting.Dictionary")
	Dim i, itemObj
	With CreateObject("MSXML2.DOMDocument")
		.Async = True
		.Load(oIni.parser("App")("LangFolder") & "\" & oIni.parser("App")("Lang") & ".xml")
		' WScript.Echo .parseError.errorCode & vbCrlf & .parseError.reason
		If .parseError.errorCode <> 0 Then oLang.Load("Locale\zh_TW.xml")
		Set itemObj = .getElementsByTagName("Item")
		For i = 0 To itemObj.Length - 1
			oLang(itemObj.Item(i).getAttribute("id")) = itemObj.Item(i).getAttribute("name")
		Next
	End With
End Sub

Sub ParseIni()
	Set oIni = New IniFile
	oIni.Read "Mailer.ini"
End Sub

Function FileExists(relativePath)
	Dim realPath
	With CreateObject("Scripting.FileSystemObject")
		realPath = .GetParentFolderName(WScript.ScriptFullName) & "\"
		realPath = realPath & relativePath
		If .FileExists(realPath) Then 
			FileExists = True
		Else 
			FileExists = False
		End If
	End With
End Function
' ---------------------------------------------------------------------------
' Subroutine:  Include
' Purpose:     Includes, or loads, other vbscript files
' Argument:    A script file name to include
' Example:     Call Include("C:\Scripts\MyScriptFile.vb")
' ---------------------------------------------------------------------------
Sub Include(strScriptName)
	With CreateObject("Scripting.FileSystemObject")
		With .OpenTextFile(strScriptName)
			ExecuteGlobal .ReadAll()
			.Close
		End With
	End With
End Sub