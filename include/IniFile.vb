'################
'#### CLASSES ###
'################
' Class 	: IE_GUI
' Purpose	: To parse an ini file
' Usage		: Set obj = New IniFile
'			  obj.Read(FilePath)

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


Class IniFile

	Public parser
	Private bufferIni, strFile, proto
	
	' read the strFile and parse it	
	Public Sub Read(sFile)
		Dim oFSO, oOTF 
		Dim sLine, sSec, aKeyValue
		strFile = sFile
		Set oFSO = CreateObject("Scripting.FileSystemObject")
		Set oOTF = oFSO.OpenTextFile(strFile)
		
		Do Until oOTF.AtEndOfStream
			sLine = oOTF.ReadLine
			bufferIni = bufferIni & sLine & vbCrlf
			sLine = Replace(sLine, vbTab, "")
			With New RegExp
				.Pattern = "^[^;\s][^;\r\n]*"		' remove the memo
				If .Test(sLine) Then
					sLine = Trim(.Execute(sLine).Item(0).Value)
					If "[" = Left(sLine, 1) Then
						sSec = RemoveQuote(sLine)
						Set proto(sSEc) = CreateObject("Scripting.Dictionary")
						Set parser(sSEc) = CreateObject("Scripting.Dictionary")
					Else
						If "" <> sLine Then
							aKeyValue = Split(sLine, "=")
							If 1 = UBound(aKeyValue) Then
								proto(sSec)(Trim(aKeyValue(0))) = RemoveQuote(Trim(aKeyValue(1)))
								parser(sSec)(Trim(aKeyValue(0))) = Scanner(aKeyValue(1))
							End If
						End If
					End If
				End If
			End With
		Loop
		oOTF.Close
		Set oFSO = Nothing
		Set oOTF = Nothing
	End Sub
	
	Public Function Write(strSec, strKey, strValue)
		Dim poSecStart, poKeyStart, poValueStart, poValueEndLine, poValueEndComment, poValueEnd
		Dim strOldValue, strNewValue, strPre, strPost
		strValue = Trim(strValue)
		poSecStart = InStr(1, bufferIni, "[" & strSec & "]", vbTextCompare)
		If poSecStart>0 Then	' Section exists
			poKeyStart = InStr(poSecStart+1, bufferIni, vbCrlf & strKey, vbTextCompare)
			If poKeyStart>0 Then	' Key exists
				poValueStart = InStr(poKeyStart+2, bufferIni, "=", vbTextCompare)
				If poValueStart > 0 Then 'Valuse exists
					poValueEndLine = InStr(poValueStart+1, bufferIni, vbCrlf, vbTextCompare)
					poValueEndComment = InStr(poValueStart+1, bufferIni, ";", vbTextCompare)
					If poValueEndComment = 0 Then poValueEndComment = Len(bufferIni) ' no comment
					'WScript.Echo poValueEndLine & vbCrlf & poValueEndComment
					If poValueEndLine>0 And poValueEndComment>0 Then
						If poValueEndLine < poValueEndComment Then 	' There is no comment on this line
							poValueEnd = poValueEndLine
						Else
							poValueEnd = poValueEndComment
						End If
						strPre = Left(bufferIni, poValueStart)
						strPost = Right(bufferIni, Len(bufferIni) - poValueEnd + 1)
						strOldValue = Mid(bufferIni, poValueStart + 1, poValueEnd - poValueStart - 1)
						strOldValue = Trim(strOldValue)
						If strOldValue = "" Then
							strNewValue = " " & strValue
						ElseIf strOldValue = """""" Then
							strNewValue = " """ & strValue & """"
						Else
							strNewValue = " " & Replace(strOldValue, proto(strSec)(strKey), strValue)
						End If
						proto(strSec)(strKey) = strValue
						parser(strSec)(strKey) = Scanner(strValue)
						bufferIni = strPre & strNewValue & strPost
						With CreateObject("Scripting.FileSystemObject")
							With .OpenTextFile(strFile, 2, True)
								.Write bufferIni
								.Close
							End With
						End With
						Write = True
						Exit Function
					End If
				End If
			End If
		End If
		Write = False
	End Function
	
	Private Function RemoveQuote(str)
		
		If ("""" = Left(str, 1) And """" = Right(str, 1)) Or _
			("[" = Left(str, 1) And "]" = Right(str, 1)) Then
			RemoveQuote = Mid(str, 2, Len(str) - 2)
		Else
			RemoveQuote = str
		End If
	End Function
	
	Private Function Scanner(str)
		str = Trim(str)
		If ("""" = Left(str, 1) And """" = Right(str, 1)) Or _
			("'" = Left(str, 1) And "'" = Right(str, 1)) Then
			Scanner = Mid(str, 2, Len(str) - 2)			' It's a string
		Else
			Scanner = str
			If LCase(str) = "true" Or LCase(str) = "on" Or LCase(str) = "yes" Then Scanner = True
			If LCase(str) = "false" Or LCase(str) = "off" Or LCase(str) = "no" Or LCase(str) = "none" Then Scanner = False
			' If str = "" Then Scanner = Null
			If IsNumeric(str) Then Scanner = CLng(str)
		End If
	End Function
	
	Private Sub Class_Initialize()
		Set parser = CreateObject("Scripting.Dictionary")
		Set proto = CreateObject("Scripting.Dictionary")
		bufferIni = ""
	End Sub

	Private Sub Class_Terminate()
		Set parser = Nothing
		Set proto = Nothing
	End Sub
	
End Class