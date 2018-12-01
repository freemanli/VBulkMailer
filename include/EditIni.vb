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
Sub EditIni()
	On Error Resume Next
	Dim sMeta, sStyle, sScript, sBody
	Dim i, key, section
	Dim flagChange : flagChange = False
	Dim sEnd : sEnd = vbCrlf
	Dim editTitle : editTitle = "SETUP"
	'WScript.Echo "Here"
	
	sMeta = "<meta name='viewport' content='width=device-width, initial-scale=1'/>" & sEnd

	sStyle = ""
	sStyle = sStyle & "<style>"  & sEnd
	sStyle = sStyle & "* {box-sizing: border-box}" & sEnd
	sStyle = sStyle & "body, html { height: 100%; margin: 0; font-family: Arial;} " & sEnd
	sStyle = sStyle & ".tablink { background-color: #555; color: white; float: left; border: none; outline: none; cursor: pointer; padding: 14px 16px; font-size: 17px; width: 20%;} " & sEnd
	sStyle = sStyle & ".tablink:hover { background-color: #777;}" & sEnd
	sStyle = sStyle & ".tabcontent { color: white; display: none; padding: 80px 20px; height: 100%;}" & sEnd
	sStyle = sStyle & ".article { width:99%;}" & sEnd
	sStyle = sStyle & ".inptext { width:90%;}" & sEnd
	sStyle = sStyle & ".tdkey { width:25%;}" & sEnd
	sStyle = sStyle & ".explain { font-size:small;width:30%;}" & sEnd
	For i = 1 To 5
		sStyle = sStyle & "#Tab" & i & " {background-color: Gray;}" & sEnd
	Next
	sStyle = sStyle & "</style>" & sEnd
	
	sScript = ""
	sScript = sScript & "<script type=""text/javascript"">" & sEnd
	sScript = sScript & "function change(sec, key) { " & sEnd
	sScript = sScript & 	"document.title='"& editTitle & "' + '_' + sec + '_' + key;" & sEnd
	sScript = sScript & 	"document.getElementById(sec + '_' + key + '_status').innerHTML = '<img src=""mail_progress.gif"" style=""width:17px"" />'; " & sEnd
	sScript = sScript & "} " & sEnd
	sScript = sScript & "function openPage(pageName, elmnt, color) { " & sEnd
						' // Hide all elements with class="tabcontent" by default */
	sScript = sScript & "var i, tabcontent, tablinks; " & sEnd
	sScript = sScript & "tabcontent = document.getElementsByClassName('tabcontent'); " & sEnd
	sScript = sScript & "for (i = 0; i < tabcontent.length; i++) {tabcontent[i].style.display = 'none';} " & sEnd
						' // Remove the background color of all tablinks/buttons
	sScript = sScript & "tablinks = document.getElementsByClassName('tablink'); " & sEnd
	sScript = sScript & "for (i = 0; i < tablinks.length; i++) { tablinks[i].style.backgroundColor = '';} " & sEnd
						' // Show the specific tab content
	sScript = sScript & "document.getElementById(pageName).style.display = 'block'; " & sEnd
						' // Add the specific color to the button used to open the tab content
	sScript = sScript & "elmnt.style.backgroundColor = color; } " & sEnd
						' // Get the element with id="TabLink1" and click on it
	sScript = sScript & "document.getElementById('TabLink1').click();" & sEnd
	sScript = sScript & "</script>"

	sBody = ""
	For i = 1 To 5
		sBody = sBody & "<button id='TabLink" & i & "' class='tablink' onclick='openPage(""Tab" & i & """, this, ""Gray"")'>" & oLang("800" & i) & "</button>" & sEnd
	Next

	' SMTP
	sBody = sBody & "<div id='Tab1' class='tabcontent'>" & sEnd
	sBody = sBody & "<table class='article'>"
	For Each key In oIni.parser("SMTP")
		sBody = sBody & "<tr>"
		sBody = sBody & "<td class='tdkey'>" & key & "</td>"
		sBody = sBody & "<td> = </td>"
		sBody = sBody & "<td><input type='"
		If key = "sendpassword" Then
			sBody = sBody & "password" 
		Else
			sBody = sBody & "text"
		End If
		sBody = sBody & "' class='inptext' id='SMTP_" & key & "' value='" & oIni.parser("SMTP")(key) & "' onchange='change(""SMTP"",""" & key & """)'/>"
		sBody = sBody & "<span id='SMTP_" & key & "_status'></span></td>"
		sBody = sBody & "<td class='explain'>" & oLang("SMTP_" & key ) & "</td>"
		sBody = sBody & "</tr>"
	Next
	sBody = sBody & "</table></div>" & sEnd

	' Sender
	sBody = sBody & "<div id='Tab2' class='tabcontent'><h3>" & oLang("8002") & "</h3>" & sEnd
	sBody = sBody & "<table class='article'>"
	For Each key In oIni.parser("Sender")
		sBody = sBody & "<tr>"
		sBody = sBody & "<td class='tdkey'>" & key & "</td>"
		sBody = sBody & "<td> = </td>"
		sBody = sBody & "<td><input type='text' class='inptext' id='Sender_" & key & "' value='" & oIni.parser("Sender")(key) & "' onchange='change(""Sender"",""" & key & """)'/>"
		sBody = sBody & "<span id='Sender_" & key & "_status'></span></td>"
		sBody = sBody & "<td class='explain'>" & oLang("Sender_" & key ) & "</td>"
		sBody = sBody & "</tr>"
	Next
	sBody = sBody & "</table><h3>" & oLang("8012") & "</h3><table class='article'>"
	For Each key In oIni.parser("Bcc")
		sBody = sBody & "<tr>"
		sBody = sBody & "<td class='tdkey'>" & key & "</td>"
		sBody = sBody & "<td> = </td>"
		sBody = sBody & "<td><input type='text' class='inptext' id='Bcc_" & key & "' value='" & oIni.parser("Bcc")(key) & "' onchange='change(""Bcc"",""" & key & """)'/>"
		sBody = sBody & "<span id='Bcc_" & key & "_status'></span></td>"
		sBody = sBody & "<td class='explain'>" & oLang("Bcc_" & key ) & "</td>"
		sBody = sBody & "</tr>"
	Next
	sBody = sBody & "</table></div>" & sEnd
	
	' Receiver
	sBody = sBody & "<div id='Tab3' class='tabcontent'><strong>" & oLang("8013") & "</strong><br/>" & sEnd
	sBody = sBody & "<table class='article'>"
	For Each key In oIni.parser("MailTo")
		sBody = sBody & "<tr>"
		sBody = sBody & "<td class='tdkey'>" & key & "</td>"
		sBody = sBody & "<td> = </td>"
		sBody = sBody & "<td><input type='text' class='inptext' id='MailTo_" & key & "' value='" & oIni.parser("MailTo")(key) & "' onchange='change(""MailTo"",""" & key & """)'/>"
		sBody = sBody & "<span id='MailTo_" & key & "_status'></span></td>"
		sBody = sBody & "<td class='explain'>" & oLang("MailTo_" & key ) & "</td>"
		sBody = sBody & "</tr>"
	Next
	sBody = sBody & "</table><strong>" & oLang("8023") & "</strong><br/><table class='article'>"
	For Each key In oIni.parser("ColumnName")
		sBody = sBody & "<tr>"
		sBody = sBody & "<td class='tdkey'>" & key & "</td>"
		sBody = sBody & "<td> = </td>"
		sBody = sBody & "<td><input type='text' class='inptext' id='ColumnName_" & key & "' value='" & oIni.parser("ColumnName")(key) & "' onchange='change(""ColumnName"",""" & key & """)'/>"
		sBody = sBody & "<span id='ColumnName_" & key & "_status'></span></td>"
		sBody = sBody & "<td class='explain'>" & oLang("ColumnName_" & key ) & "</td>"
		sBody = sBody & "</tr>"
	Next
	sBody = sBody & "</table></div>" & sEnd
	
	' Content
	sBody = sBody & "<div id='Tab4' class='tabcontent'>" & sEnd
	sBody = sBody & "<ul style='font-size:small'><li><strong>" & oLang("8014") & "</strong><br/>" & _
					oLang("8024") & "</li><li>" & _
					oLang("8034") & "</li><li>" & _
					oLang("8044") & "</li><li>" & _
					oLang("8054") & "</li></ul>" & sEnd
	sBody = sBody & "<table class='article'>" & sEnd
	For Each key In oIni.parser("Letter")
		sBody = sBody & "<tr>"
		sBody = sBody & "<td class='tdkey'>" & key & "</td>"
		sBody = sBody & "<td> = </td>"
		sBody = sBody & "<td><input type='text' class='inptext' id='Letter_" & key & "' value='" & oIni.parser("Letter")(key) & "' onchange='change(""Letter"",""" & key & """)'/>"
		sBody = sBody & "<span id='Letter_" & key & "_status'></span></td>"
		sBody = sBody & "<td class='explain'>" & oLang("Letter_" & key ) & "</td>"
		sBody = sBody & "</tr>"
	Next
	sBody = sBody & "</table></div>" & sEnd
	' Attachment
	sBody = sBody & "<div id='Tab5' class='tabcontent'><ul style='font-size:small'><li><strong>" & _
					oLang("8035") & "</strong></li><li>" & _
					oLang("8045") & "</li><li>" & _
					oLang("8055") & "</li></ul>" & sEnd
	sBody = sBody & "<table class='article'>" & sEnd
	sBody = sBody & "<tr><th>" & _
					oLang("8015") & "</th><th>" & _
					oLang("8025") & "</th></tr>" & sEnd
	For i = 0 To 4
		sBody = sBody & "<tr>"
		sBody = sBody & "<td><input type='text' class='inptext' id='Attachements_File" & i & "' value='" & oIni.parser("Attachements")("File"&i) & "' onchange='change(""Attachements"",""File" & i & """)'/>"
		sBody = sBody & "<span id='Attachements_File" & i & "_status'></span></td>"
		sBody = sBody & "<td><input type='text' class='inptext' id='HisAttachements_File" & i & "' value='" & oIni.parser("HisAttachements")("File"&i) & "' onchange='change(""HisAttachements"",""File" & i & """)'/>"
		sBody = sBody & "<span id='HisAttachements_File" & i & "_status'></span></td>"
		sBody = sBody & "</tr>"
	Next
	sBody = sBody & "</table>"
	sBody = sBody & "<p><button name='close' type='button' onclick='" 
	sBody = sBody & "document.title=""" & editTitle & "_close"";' style='font-weight: bold;float:right;'>" & oLang("8006") & "</button></p>"
	sBody = sBody & "</div>" & sEnd
	
	Dim editGUI : Set editGUI = New IE_GUI
	With editGUI
		.window("type") = IE_GUI_HTML
		.dialog("title") = editTitle
		.dialog("Width") = 640
		.dialog("Height") = 480

		
		.dialog("head") = sMeta & sStyle
 		.dialog("body") = sBody & sScript
 		.Show "editGUI"
		.Scroll = False
		
		Dim newTitle, iniSecKey, flagFile, flagIniWrite, checkFile, checkNo, newValue
		While Not IsEmpty(.Title)
			newTitle = .Title
			If newTitle <> editTitle Then
				flagFile = True
				flagIniWrite = True
				newTitle = Replace(newTitle, editTitle & "_", "")
				If newTitle = "close" Then
					.Close
				Else
					iniSecKey = Split(newTitle, "_")
					If IsArray(iniSecKey) And UBound(iniSecKey)=1 Then
						newValue = .GetElementByID(newTitle).value
						If InStr(iniSecKey(1), "File")>0 And Trim(newValue) <> "" Then	' check file exist
							checkFile = fso.GetParentFolderName(WScript.ScriptFullName) & "\"
							checkFile = checkFile & oIni.parser("App")("MailFolder") & "\"
							If iniSecKey(0) = "HisAttachements" Then
								checkNo = oIni.parser("MailTo")("StartNo")
								If checkNo = "" Then checkNo = 1
								checkFile = checkFile & Replace(newValue, "%s", checkNo)
							Else
								checkFile = checkFile & newValue
							End If
							If Not fso.FileExists(checkFile) Then
								flagFile = False
								'WScript.Echo checkFile
								.GetElementByID(newTitle).value = oIni.parser(iniSecKey(0))(iniSecKey(1))
							End If
						End If
						If flagFile Then
							If Not oIni.Write(iniSecKey(0), iniSecKey(1), .GetElementByID(newTitle).value) Then
								flagIniWrite = False
							End If
						End If
						If flagFile And flagIniWrite Then
							.GetElementByID(newTitle & "_status").innerHTML = "<img src=""mail_okay.png"" style=""width:17px;"" />"
							flagChange = True
						Else
							.GetElementByID(newTitle & "_status").innerHTML = "<img src=""mail_error.png"" style=""width:17px;"" />"
						End If
					End If
				End If
				.Title = editTitle
			End If
			WScript.Sleep 200
		WEnd
	End With
	WScript.Sleep 200
	Set editGUI = Nothing

	Dim flagInitial : flagInitial = False
	While Not flagInitial
		flagInitial = Initialize()
		If Not flagInitial Then 
			If Popup(oLang("1005"), , oLang("1002"), vbYesNo) = vbYes Then
				Call EditIni()
			Else
				Call EndProcess()
			End If
		End If
	WEnd
End Sub
