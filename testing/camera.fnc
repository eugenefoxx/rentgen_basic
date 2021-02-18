'<--FNC-Version> 4 (Inserted by FNC Editor - DO NOT EDIT!)

Sub Main()

Dim IPCurrent As FNCImageProcess
	Set IPCurrent = IPEditor.GetIPFromName(IPEditor.ActiveIP)
	If IPCurrent.InputState = fncISCamera Then
		MsgBox "Camera is active"
	End If

End Sub
