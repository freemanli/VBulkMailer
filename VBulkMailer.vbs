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
Include("include\class_IniFile.vb")
Include("include\class_IE_GUI.vb")
Include("include\class_xlsreader.vb")
Include("include\class_ExtMessage.vb")
Include("include\Popup.vb")
Include("include\Proc_EditIni.vb")
Include("include\Proc_Mail.vb")
Const WAIT_TIME = 5000
Const INI_FILE = "VBulkMailer.ini"
' Declare Global Variable
Dim oIni, oLang, oExcel, oMsg, sLetter, fso, procGUI, scriptPath

Call Main

Sub Main()
	Dim execStep, sBody, i
	Dim mailerTitle : mailerTitle = "VBulkMailer"
	
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

Function Initialize()
	Dim initialProgess : Set initialProgess = New IE_GUI
	Set fso = CreateObject("Scripting.FileSystemObject")
	' copy image files
	With fso
		Dim sourcePath : sourcePath = .GetParentFolderName(WScript.ScriptFullName)& "\"
		scriptPath = sourcePath
		Dim destinationPath : destinationPath = .GetSpecialFolder(2) & "\"
		.CopyFile sourcePath & "mail_error.png", destinationPath & "mail_error.png"
		.CopyFile sourcePath & "mail_okay.png", destinationPath & "mail_okay.png"
		.CopyFile sourcePath & "mail_progress.gif", destinationPath & "mail_progress.gif"
		If Not .FolderExists(scriptPath & "\log") Then .CreateFolder(scriptPath & "\log")
		If Not .FolderExists(scriptPath & "\mail") Then .CreateFolder(scriptPath & "\mail")
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
	Dim pos, pStart, pEnd, sCharSet
	Dim letterPath : letterPath = oIni.parser("App")("MailFolder") & "\" & oIni.parser("Letter")("File")
	'WScript.Echo "Here !"
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
	If FileExists(INI_FILE) Then
		oIni.Read INI_FILE
	ElseIf FileExists(INI_FILE & ".default") Then
		fso.CopyFile scriptPath & INI_FILE & ".default", scriptPath & INI_FILE
		oIni.Read INI_FILE
	Else
		WScript.Echo Replace("Can not Find %1. VBulkMailer System Terminated!", "%1", INI_FILE)
		Call EndProcess()
	End If
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