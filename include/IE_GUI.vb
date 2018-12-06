'################
'#### CLASSES ###
'################
' Class 	: IE_GUI
' Purpose	: To create a GUI window based on an HTML page displayed in an Internet Explorer window.
' 			  There are 2 kinds of GUI : progress bar, html
' Usage		: Set obj = New IE_GUI

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

' ================================ Example 1 : The usage of progress bar ================================
' 	Option Explicit
' 	Dim progress : Set progress = New IE_GUI
' 	With progress
' 		.window("type") = IE_GUI_PROGRESS_BAR
' 		.dialog("title") = "MyTitle"
' 		.dialog("head") = "MyMainText 主旨"
' 		.dialog("body") = "MySubText"
' 		.Show "progress"					' this argument must be the same name as the variable name of new class
' 	End With
' 	
' 	Dim i : i = 0
' 	While progress.window("exist") And i <= 100
' 	 	WScript.Sleep(1000)
' 	 	i = i + 5
' 	 	progress.dialog("head") = "MyMainText 主旨" & " : " & i
' 	 	progress.dialog("body") = i & "% proceed! Now we test the very long sentence and see what will show on the dialog box."
' 	 	progress.SetPct(i)
' 	WEnd
' 	progress.Close
' 	If i<100 Then
' 		WScript.Echo "You terminate it at i = " & i
' 	Else
' 		WScript.Echo "bye! bye! Example 1"
' 	End If
' 	Set progress = Nothing
' 	' ================================ End of Example 1 ================================
' 	
' 	' ================================ Example 2 : The usage of html ================================
' 	Dim myVariableTitle : myVariableTitle = "My Letter"
' 	Dim myLetter : Set myLetter = New IE_GUI
' 	With myLetter
' 		.window("type") = IE_GUI_HTML
' 		.dialog("title") = myVariableTitle
' 		.dialog("Width") = 960
' 		.dialog("Height") = 300
' 		.dialog("head") = ""
' 		.dialog("body") = 	"<script>function myFunction() {document.title=document.title + '_a';}</script>" & _ 
' 							"<button onclick='myFunction()'>Click me</button>" & _
' 							"<div id='demo'></div>" & _
' 							"<script>document.getElementById('demo').innerHTML = Date();</script>"
' 		.Show "myLetter"			' this argument must be the same name as the variable name of new class
' 	End With
' 	
' 	Dim j : j = 0
' 	Dim echoStr : echoStr =""
' 	While myLetter.window("exist") And j <= 100
' 		WScript.Sleep(200)
' 		j = j + 1
' 		If myLetter.Title <> myVariableTitle And myLetter.Title<>"" Then
' 			echoStr = "You change the title " & "<br/>" & "from : " & myVariableTitle & "<br/>"
' 			myVariableTitle = myLetter.Title
' 			echoStr = echoStr & "to : " & myVariableTitle & vbCrlf
' 			myLetter.GetElementByID("demo").innerHTML = echoStr
' 		End If
' 	WEnd
' 	myLetter.close
' 	If j<100 Then
' 		WScript.Echo "You terminate it at j = " & j
' 	Else
' 		WScript.Echo "bye! bye! Example 2"
' 	End If
' 	Set myLetter = Nothing
' 	' ================================ End of Example 2 ================================

' #############################################################################################################################################################

' These constants are used to create the progress bar for IE < 10
Const IE_GUI_SOLID_BLOCK_CHARACTER = "■"
Const IE_GUI_EMPTY_BLOCK_CHARACTER = "□"
Const IE_GUI_SOLID_BLOCK_COLOUR = "#ffcc33;"
Const IE_GUI_EMPTY_BLOCK_COLOUR = "#666666;"
Const IE_GUI_BLOCKSCOUNT = 28
' These constants are used to define the type of GUI
Const IE_GUI_PROGRESS_BAR = 1
Const IE_GUI_HTML = 2
Const IE_GUI_FONT_FAMILY = "微軟正黑體, Arial, Helvetica"


Class IE_GUI

	Private tempFile, oGUI

	' Properties of the class 
	Public dialog, window
	
	Public Property Let Title(stitle)
		On Error Resume Next
		oGUI.Document.Title = stitle
		If Err.Number <> 0 Then Close
	End Property
	
	Public Property Get Title()
		On Error Resume Next
		Title = oGUI.Document.Title
		If Err.Number <> 0 Then Close
	End Property
	
	Public Property Let Visible(flagVisible)
		On Error Resume Next
		oGUI.Visible = flagVisible
		If Err.Number <> 0 Then Close
	End Property
	
	Public Property Let Scroll(flagScroll)
		On Error Resume Next
		If flagScroll Then 
			oGUI.Document.Body.Scroll = "yes"
		Else
			oGUI.Document.Body.Scroll = "no"
		End If
		If Err.Number <> 0 Then Close
	End Property
	
	' Methods of the class 
	Public Function GetElementByID(elementID)
		On Error Resume Next
		Set GetElementByID = oGUI.Document.GetElementByID(elementID)
		If Err.Number <> 0 Then Close
	End Function
	
	Public Sub Activate(appName)
		On Error Resume Next
		Dim objItem, lastItem
		'Dim strDebug : strDebug = ""
		WScript.Sleep 200
		With GetObject("winmgmts:\\.\root\cimv2")
			For Each objItem In .ExecQuery("Select * From Win32_Process where name='" & appName & "'")
				'strDebug = strDebug & objItem.ProcessId & " -> " & objItem.ParentProcessId & vbCrlf
				Set lastItem = objItem
			Next
			With CreateObject("WScript.Shell")
				.AppActivate(lastItem.ParentProcessId)
				' .AppActivate(lastItem.ProcessId)
				.SendKeys "{TAB}"
			End With
			'WScript.Echo strDebug
		End With
	End Sub
	
	Public Sub Show(GUIName)
		Select case window("type") 
			case IE_GUI_PROGRESS_BAR
				dialog("innerHTML") = "<html><head></head><body>" & _
							"<div id='maintext' style='margin:0px 10px;font-weight:bold;font-family: " & IE_GUI_FONT_FAMILY & ";'>" & dialog("head") & "</div>" & _
							"<div id='progress' style='margin:10px 10px 0px;'>" & ProgressHTML(0) & "</div>" & _
							"<div id='subtext' style='margin:5px 10px;font-size:small;font-family: " & IE_GUI_FONT_FAMILY & ";'>" & dialog("body") & "</div></body></html>"
				dialog("scroll") = "no"
				window("navigate") = "about:blank"
			case IE_GUI_HTML
				dialog("innerHTML") = "<!DOCTYPE html>" & _
							"<html>" & _
							"<head>" & dialog("head") & "</head>" & _
							"<body>" & dialog("body") & "</body>" & _
							"</html>"
				dialog("scroll") = "yes"
				With CreateObject("Scripting.FileSystemObject")
					Dim tempName : tempName = .GetTempName
					tempFile = .GetSpecialFolder(2) & "\" & tempName & ".html"
					With .CreateTextFile(tempFile, True)
						.Write dialog("innerHTML")
						.Close
					End With
				End With
				dialog("scroll") = "yes"
				window("navigate") = "file:///" & tempFile
			case Else
				WScript.Quit
		End Select
	
		Set oGUI = WScript.CreateObject("InternetExplorer.Application", GUIName & "_")
		' ExecuteGlobal "Sub " & GUIName & "_OnQuit() : " & GUIName & ".Close : End Sub"
		' =============================== Unable to get object of InternetExplorer.Application under Windows 10, There is no OnQuit event.
		SetCommonValue(oGUI)
		Dim hWnd : hWnd = oGUI.HWND 	
		' receives the window handle
		oGUI.Navigate2 window("navigate")
		If window("type") = IE_GUI_PROGRESS_BAR Then oGUI.Document.Body.InnerHTML = dialog("innerHTML")
		WScript.Sleep 100
		Dim appWindow
		For Each appWindow In CreateObject("Shell.application").Windows
			If hWnd = appWindow.HWND Then
				Set oGUI = appWindow
				Exit For
			End If
		Next
		' ===============================
		With oGUI
			Do While .ReadyState <> 4
				Wscript.sleep 10
			Loop
			.Document.Title = dialog("title")
			.Document.Body.Scroll = dialog("scroll")
			.Visible = True
		End With
		window("exist") = True
		Activate "iexplore.exe"
	End Sub
	
	Private Sub SetCommonValue(oIE)
		With oIE
			.Width = dialog("Width")
			.Height = dialog("Height")
			If dialog("Left") = -1 Then dialog("Left") = (window("ScreenWidth") - dialog("Width"))/2
			.Left = dialog("Left")
			If dialog("Top") = -1 Then dialog("Top") = (window("ScreenHeight") - dialog("Height") - 40)/2
			.Top = dialog("Top")
			.Toolbar = dialog("Toolbar")
			.StatusBar = dialog("StatusBar")
			.Resizable = dialog("Resizable")
		End With
	End Sub
	
	' ================================ ProgressBar Start ================================
	Public Sub SetPct(vPct)
		On Error Resume Next
		If window("exist") And window("type") = IE_GUI_PROGRESS_BAR Then 
			GetElementByID("maintext").innerHTML = dialog("head")
			GetElementByID("progress").innerHTML = ProgressHTML(vPct)
			GetElementByID("subtext").innerHTML = dialog("body")
		End If
	End Sub
	
	Private Function ProgressHTML(pct)
		Dim progStr : progStr = ""
		
		' For IE version < 10 Calculate how many blocks we need to create
		Dim solidBlocks, emptyBlocks, block
		solidBlocks = round(pct / 100 * IE_GUI_BLOCKSCOUNT)
		emptyBlocks = IE_GUI_BLOCKSCOUNT - solidBlocks
		block = 0
		progStr = "<!--[if IE]>"				' Hack Words Start, For IE version < 10, Internet Explorer conditional comment
		progStr = progStr & "<span style='font-family:courier; color:'" & IE_GUI_SOLID_BLOCK_COLOUR & "'>"
		While block < solidBlocks
			progStr = progStr & IE_GUI_SOLID_BLOCK_CHARACTER
			block = block + 1
		Wend
		progStr = progStr &  "</span><span style='font-family:courier;color:" & IE_GUI_EMPTY_BLOCK_COLOUR & ";'>"
		block = 0
		While block < emptyBlocks
			progStr = progStr & IE_GUI_EMPTY_BLOCK_CHARACTER
			block = block + 1
		Wend
		progStr = progStr &  "</span>"
		progStr = progStr &  "<![endif]-->" 	' Hack Words End, Internet Explorer conditional comment
		
		progStr = progStr & "<progress  value='" & pct & "' max='100' style='width:100%;height:15px;'></progress>"
		ProgressHTML = progStr
	End Function
	' ================================ ProgressBar End ================================
	
	Public Sub Close()
		On Error Resume Next
		If window("exist") Then
			' CreateObject("WScript.Shell").Run "taskkill /im iexplore.exe", 0, True
			oGUI.Quit
			Set oGUI = Nothing
			CreateObject("Scripting.FileSystemObject").DeleteFile tempFile, True
			window("exist") = False
		End If
	End Sub
	
	Private Sub Class_Initialize()
		Set dialog = CreateObject("Scripting.Dictionary")
		Set window = CreateObject("Scripting.Dictionary")
		' Get Screen Resolution
		With GetObject("winmgmts:\\.\root\cimv2")
			Dim objItem 
			For Each objItem in .ExecQuery("Select * from Win32_DesktopMonitor",,48)
				window("ScreenWidth") = objItem.ScreenWidth
				window("ScreenHeight") = objItem.ScreenHeight
			Next
			' Not PNP Screen
			If IsNull(window("ScreenWidth")) Then window("ScreenWidth") = 1024
			If IsNull(window("ScreenHeight")) Then window("ScreenHeight") = 768
		End With
		' Set Default Value
		dialog("Left") = -1	' Centered
		dialog("Top") = -1 	' Centered
		dialog("Width") = 320
		dialog("Height") = 150
		dialog("Toolbar") = False
		dialog("StatusBar") = False
		dialog("Resizable") = False
	End Sub

	Private Sub Class_Terminate()
		On Error Resume Next
		oGUI.Quit
		Set oGUI = Nothing
		Set dialog = Nothing
		Set window = Nothing
	End Sub
	
End Class