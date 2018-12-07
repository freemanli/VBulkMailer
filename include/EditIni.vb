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
	Dim sEnd : sEnd = ""
	Dim editTitle : editTitle = "SETUP"
	'WScript.Echo "Here"
	
	sMeta = "<meta name='viewport' content='width=device-width, initial-scale=1'/>" & sEnd
	sMeta = sMeta & "<meta http-equiv='content-type' content='text/html; charset=big5'>"
	'sMeta = sMeta & "<meta http-equiv='content-type' content='text/html; charset=utf-8'>"

	sStyle = ""
	sStyle = sStyle & "<style>"  & sEnd
	sStyle = sStyle & "* {box-sizing: border-box}" & sEnd
	sStyle = sStyle & "body, html { height: 100%; margin: 0; font-family: Arial;} " & sEnd
	sStyle = sStyle & ".tablink { background-color: #555; color: white; float: left; border: none; outline: none; cursor: pointer; padding: 14px 16px; font-size: 17px; width: 20%;} " & sEnd
	sStyle = sStyle & ".tablink:hover { background-color: #777;}" & sEnd
	sStyle = sStyle & ".tabcontent { color: white; display: none; padding: 80px 20px; height: 100%;}" & sEnd
	sStyle = sStyle & ".article { width:99%;}" & sEnd
	sStyle = sStyle & ".secTitle { font-weight: bold;color:blue;}" & sEnd
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
	sBody = sBody & "<div id='Tab2' class='tabcontent'>"
	Dim Sender(2)
	Sender(0) = "Sender"
	Sender(1) = "Bcc"
	Sender(2) = "TestTo"
	For i = 0 To 2
		sBody = sBody & "<span class='secTitle'>" & oLang(Sender(i)) & "</span><br/>" & sEnd
		sBody = sBody & "<table class='article'>"
		For Each key In oIni.parser(Sender(i))
			sBody = sBody & "<tr>"
			sBody = sBody & "<td class='tdkey'>" & key & "</td>"
			sBody = sBody & "<td> = </td>"
			sBody = sBody & "<td><input type='text' class='inptext' id='" & Sender(i) & "_" & key & "' value='" & oIni.parser(Sender(i))(key) & "' onchange='change(""" & Sender(i) & """,""" & key & """)'/>"
			sBody = sBody & "<span id='" & Sender(i) & "_" & key & "_status'></span></td>"
			sBody = sBody & "<td class='explain'>" & oLang(Sender(i) & "_" & key ) & "</td>"
			sBody = sBody & "</tr>"
		Next
		sBody = sBody & "</table>" & sEnd
	Next
	sBody = sBody & "</div>" & sEnd
	
	' Receiver
	sBody = sBody & "<div id='Tab3' class='tabcontent'><span class='secTitle'>" & oLang("8013") & "</span><br/>" & sEnd
	sBody = sBody & "<table class='article'>"
	For Each key In oIni.parser("MailTo")
		sBody = sBody & "<tr>"
		sBody = sBody & "<td class='tdkey'>" & key & "</td>"
		sBody = sBody & "<td> = </td><td>"
		If key = "File" Then
			sBody = sBody & "<input type='text'  class='inptext' id='MailTo_" & key & "' value='"
			If FileExists(oIni.parser("App")("MailFolder") & "\" & oIni.parser("MailTo")(key)) Then
				sBody = sBody & oIni.parser("MailTo")(key) 
			Else 	
				Call oIni.Write("MailTo", "File", "")
			End If
			sBody = sBody & "' disabled><br/>"
			sBody = sBody & "<input type='file' id='MailTo_" & key & "_Path' value='' " 
			sBody = sBody & "accept='.csv, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/vnd.ms-excel' " 
			sBody = sBody & "onchange='change(""MailTo"",""" & key & """)'/>"
		Else
			sBody = sBody & "<input type='text' class='inptext' id='MailTo_" & key & "' value='" & oIni.parser("MailTo")(key) & "' onchange='change(""MailTo"",""" & key & """)'/>"
		End If
		sBody = sBody & "<span id='MailTo_" & key & "_status'></span>"
		
		sBody = sBody & "</td><td class='explain'>" & oLang("MailTo_" & key ) & "</td>"
		sBody = sBody & "</tr>"
	Next
	sBody = sBody & "</table><span class='secTitle'>" & oLang("8023") & "</span><br/><table class='article'>"
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
		sBody = sBody & "<td> = </td><td>"
		If key = "File" Then
			sBody = sBody & "<input type='text'  class='inptext' id='Letter_" & key & "' value='"
			If FileExists(oIni.parser("App")("MailFolder") & "\" & oIni.parser("Letter")(key)) Then
				sBody = sBody & oIni.parser("Letter")(key) 
			Else 	
				Call oIni.Write("Letter", "File", "")
			End If
			sBody = sBody & "' disabled><br/>"
			sBody = sBody & "<input type='file' id='Letter_" & key & "_Path' value='' " 
			sBody = sBody & "accept='.txt, .htm, .html' " 
			sBody = sBody & "onchange='change(""Letter"",""" & key & """)'/>"
		Else
			sBody = sBody & "<input type='text' class='inptext' id='Letter_" & key & "' value='" & oIni.parser("Letter")(key) & "' onchange='change(""Letter"",""" & key & """)'/>"
		End If
		sBody = sBody & "<span id='Letter_" & key & "_status'></span>"
		sBody = sBody & "</td><td class='explain'>" & oLang("Letter_" & key ) & "</td>"
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
		sBody = sBody & "<tr><td>"

		sBody = sBody & "<input type='text'  class='inptext' id='Attachements_Attach" & i & "' value='"
		If FileExists(oIni.parser("App")("MailFolder") & "\" & oIni.parser("Attachements")("Attach" & i)) Then
			sBody = sBody & oIni.parser("Attachements")("Attach" & i) 
		Else 	
			Call oIni.Write("Attachements", "Attach" & i, "")
		End If
		sBody = sBody & "' "
'		sBody = sBody & "<input type='file' id='Attachements_" & "File" & i & "_Path' value='' " 
'		sBody = sBody & "accept='text/plain, text/html, .htm' " 
		sBody = sBody & "onchange='change(""Attachements"",""" & "Attach" & i & """)'/>"
		
'		sBody = sBody & "<input type='text' class='inptext' id='Attachements_File" & i & "' value='" & oIni.parser("Attachements")("File"&i) & "' onchange='change(""Attachements"",""File" & i & """)'/>"
		sBody = sBody & "<span id='Attachements_Attach" & i & "_status'></span>"

		sBody = sBody & "</td><td><input type='text' class='inptext' id='HisAttachements_Attach" & i & "' value='" & oIni.parser("HisAttachements")("Attach"&i) & "' onchange='change(""HisAttachements"",""Attach" & i & """)'/>"
		sBody = sBody & "<span id='HisAttachements_Attach" & i & "_status'></span></td>"
		sBody = sBody & "</tr>"
	Next
	sBody = sBody & "</table><p>"
	sBody = sBody & "<button name='exit' type='button' onclick='" 
	sBody = sBody & "document.title=""" & editTitle & "_exit"";' style='font-weight: bold;float:right; margin:5px 10px; color:red; '>" & oLang("7000") & "</button>"
	sBody = sBody & "<button name='close' type='button' onclick='" 
	sBody = sBody & "document.title=""" & editTitle & "_close"";' style='font-weight: bold;float:right; margin:5px 10px;'>" & oLang("8006") & "</button>"
	sBody = sBody & "</p></div>" & sEnd
	
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
		
		Dim newTitle, iniSecKey, flagFile, flagIniWrite, checkFile, checkNo, newValue, sourcePath, targetPath, wsName
		
		While Not IsEmpty(.Title)
			newTitle = .Title
			If newTitle <> editTitle Then
				flagFile = True
				flagIniWrite = True
				newTitle = Replace(newTitle, editTitle & "_", "")
				
				If newTitle = "close" Then
					.Close
				ElseIf newTitle = "exit" Then
					.Close
					Call EndProcess()
				Else
					iniSecKey = Split(newTitle, "_")
					' Make Sure it is a changing event
					If IsArray(iniSecKey) And UBound(iniSecKey)=1 Then
						newValue = .GetElementByID(newTitle).value
						If ( InStr(iniSecKey(1), "File")>0 Or InStr(iniSecKey(1), "Attach")>0 ) _
							And Trim(newValue) <> "" Then		
							' Files And Attachements
							targetPath = fso.GetParentFolderName(WScript.ScriptFullName) & "\"
							targetPath = targetPath & oIni.parser("App")("MailFolder") & "\"
							
							If InStr(iniSecKey(1), "File")>0 Then	' File
								sourcePath = .GetElementByID(newTitle & "_Path").value
								targetPath = targetPath & Mid(sourcePath, InStrRev(sourcePath,"\")+1)
								' Copy File
								fso.CopyFile sourcePath, targetPath
								.GetElementByID(newTitle).value = Mid(sourcePath, InStrRev(sourcePath,"\")+1)
								' Now to modify MailTo_Worksheet and Letter_Format
								If iniSecKey(0) = "Letter" Then
									If LCase(Right(targetPath, 3)) = "txt" Then
										.GetElementByID("Letter_Format").value = "TEXT"
										Call oIni.Write("Letter", "Format", "TEXT")
									Else
										.GetElementByID("Letter_Format").value = "HTML"
										Call oIni.Write("Letter", "Format", "HTML")
									
									End If
								Else	' EXCEL
									With CreateObject("Excel.Application")
										' WScript.Echo  targetPath
										.Workbooks.Open targetPath
										wsName = .Worksheets(1).Name
										.Workbooks.Close False
										.Workbooks.Quit
									End With
									.GetElementByID("MailTo_Worksheet").value = wsName
									Call oIni.Write("MailTo", "Worksheet", wsName)
								End If
								
							Else									' Attachemensts
								If iniSecKey(0) = "HisAttachements" Then
									checkNo = oIni.parser("MailTo")("StartNo")
									If checkNo = "" Then checkNo = 1
									targetPath = targetPath & Replace(newValue, "%s", checkNo)
								Else
									targetPath = targetPath & newValue
								End If
							End If
							
							If Not fso.FileExists(targetPath) Then
								flagFile = False
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
