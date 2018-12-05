'################
'#### CLASSES ###
'################
'Create an instance of this class to read an excel file

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


Class ExcelFile
	
	Public columnName, rowData
	
	Private progressBar, pct
	
	Public Sub SetProgress(prog, pc)
		Set progressBar = prog
		pct = pc
	End Sub
	
	Private Sub SetPct(pc)
		progressBar.SetPct(pc)
	End Sub
	
	
	Public Function Read(excelStr, sheetStr, startNo, endNo)
		On Error Resume Next
		Dim excelObj, sheetObj, usedRowsCount, usedColsCount
		Dim rowNo, colNo
		Dim flagRead : flagRead = True
		Set excelObj = CreateObject("Excel.Application")

		' Don't display any alert messages
		excelObj.DisplayAlerts = 0  
		
		With CreateObject("Scripting.FileSystemObject")
			Dim bookPath : bookPath = .GetParentFolderName(WScript.ScriptFullName)& "\" & excelStr			
		End With
		excelObj.Workbooks.Open bookPath
		If Err.Number <> 0 Then 
			flagRead = False
			Exit Function
		End If
		SetPct(pct+(100-pct)*.1)
'		excelObj.Visible = True

		If sheetStr = "" Then
			' read the first sheet if no other assignation
			Set sheetObj = excelObj.ActiveWorkbook.Worksheets(1)
			If Err.Number <> 0 Then 
				flagRead = False
				Exit Function
			End If
		Else
			Set sheetObj = excelObj.Worksheets(sheetStr)
			If Err.Number <> 0 Then 
				flagRead = False
				Exit Function
			End If
		End If
		SetPct(pct+(100-pct)*.2)

		' Get the number of used rows
		usedRowsCount = sheetObj.UsedRange.Rows.Count
		' Get the number of used columns
		usedColsCount = sheetObj.UsedRange.Columns.Count
		
		' read the column name
		For colNo = 1 To usedColsCount
			columnName(colNo) = sheetObj.Cells(1, colNo).Value
		Next
		SetPct(pct+(100-pct)*.3)
		' read the row data
		If startNo = "" Then startNo = 1
		If endNo = "" Then endNo = usedRowsCount
		For rowNo = startNo To endNo
			Set rowData(rowNo) = CreateObject("Scripting.Dictionary")
			For colNo = 1 To usedColsCount
				rowData(rowNo)(columnName(colNo)) = sheetObj.Cells(rowNo + 1, colNo).Value
			Next
			If (rowNo Mod 10) = 0 Then SetPct(pct+(100-pct)*(.3 + .7*rowNo/endNo))
		Next
		
'		WScript.Echo rowData(1)("¹q¤l¶l¥ó")
		Set sheetStr = Nothing
		excelObj.Quit
		Set excelObj = Nothing
		Read = flagRead
	End Function
	
	Private Sub Class_Initialize()
		Set columnName = CreateObject("Scripting.Dictionary")
		Set rowData = CreateObject("Scripting.Dictionary")
	End Sub

	Private Sub Class_Terminate()
		Set columnName = Nothing
		Set rowData = Nothing
	End Sub

End Class
