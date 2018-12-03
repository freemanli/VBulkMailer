' ---------------------------------------------------------------------------
' Function:	Popup
' Purpose:	Message Box
' Usage:	Popup(strText, nSecondsToWait, strTitle, nType)
' Argument:	strText			String value containing the text you want to appear in the pop-up message box.
'			nSecondsToWait	Numeric value indicating the maximum length of time (in seconds) you want the pop-up message box displayed.
'			strTitle		String value containing the text you want to appear as the title of the pop-up message box.
'			nType			Numeric value indicating the type of buttons and icons you want in the pop-up message box. These determine how the message box is used.
'				Button Types
'				Value	Constant 			Description 	https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/msgbox-constants
'				-------------------------------------------------------------------------------------
'				0		vbOKOnly			Show OK button.
'				1		vbOKCancel			Show OK and Cancel buttons.
'				2		vbAbortRetryIgnore	Show Abort, Retry, and Ignore buttons.
'				3		vbYesNoCancel		Show Yes, No, and Cancel buttons.
'				4		vbYesNo				Show Yes and No buttons.
'				5		vbRetryCancel		Show Retry and Cancel buttons.
'				Icon Types
'				Value	Constant			Description
'				--------------------------------------------------------------------------------------
'				16		vbCritical			Show "Stop Mark" icon.
'				32		vbQuestion			Show "Question Mark" icon.
'				48		vbExclamation		Show "Exclamation Mark" icon.
'				64		vbInformation		Show "Information Mark" icon.
'				4096	vbSystemModal 		System modal message box, Always On Top
'				16384	vbMsgBoxHelpButton 	Adds Help button to the message box
'				65536	VbMsgBoxSetForeground 		Specifies the message box window as the foreground window
'				524288	vbMsgBoxRight 		Text is right aligned
'				1048576	vbMsgBoxRtlReading	Specifies text should appear as right-to-left reading on Hebrew and Arabic systems
' Return:	Value	Constant		Description
'			------------------------------------------------------------------------------
'			1		vbOK			OK button
'			2		vbCancel		Cancel button
'			3		vbAbort			Abort button
'			4		vbRetry			Retry button
'			5		vbIgnore		Ignore button
'			6		vbYes			Yes button
'			7		vbNo			No button
'	*		-1		vbTimeUp		If the user does not click a button before nSecondsToWait seconds
' ===========================================================================================
Const vbTimeUp	= -1

Function Popup(strText, nSecondsToWait, strTitle, nType)
	Dim WshShell
	Set WshShell = WScript.CreateObject("WScript.Shell")
	'Popup = WshShell.Popup(strText, nSecondsToWait, strTitle, nType + vbSystemModal)
	Popup = WshShell.Popup(strText, nSecondsToWait, strTitle, nType)
End Function
