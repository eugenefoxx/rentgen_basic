Sub Main   '  ListBox Dialog Item Definition'
    Dim lists$()
    Dim tests$()
    Dim IPLive As FNCImageProcess
    Dim IPCurrent As FNCImageProcess
    Dim i As Integer

'	IPEditor.GetIPFromName(IPEditor.ActiveIP).ActivateCamera

'	IPEditor.GetIPFromName(IPEditor.ActiveIP).Enable(False)

	lists$ = Get_DirS("C:\WWTEMP\")

    Do While Fol <> ""
        lists$() = Dir()
    Loop
'    lists$(0) = "List 0"
'    lists$(1) = "List 1"
'    lists$(2) = "List 2"
'    lists$(3) = "List 3"
'    lists$(4) = Fol

    Begin Dialog UserDialog 600,220
        Text 10,10,18000,15,"Please push the Cancel button"
        ComboBox 10,25,18000,160,lists$(),.list$
        OKButton 10,190,40,20
        CancelButton 110,190,60,20
    End Dialog
    Dim dlg As UserDialog
 '   dlg.list = 2
    Dialog dlg ' show dialog (wait for cancel)
    Debug.Print dlg.list$
    L = dlg.list$
 '   MsgBox L

 	tests$ = Get_DirS(L & "\")

	Do While Fol <> ""
        tests$() = Dir()
    Loop

	Begin Dialog UserDialog 600,220
        Text 10,10,1800,15,"Please push the Cancel button"
        ComboBox 10,25,1800,160,tests$(),.test$
        OKButton 40,190,40,20
        CancelButton 110,190,60,20
    End Dialog
    Dim dlgS As UserDialog
 '   dlg.list = 2
 	Line1:
    Dialog dlgS ' show dialog (wait for cancel)
    Debug.Print dlgS.test$
    S = dlgS.test$ & "\"
'	MsgBox S
	For i = 0 To 100
'	If (MsgBox("Save image?",vbYesNo) = vbNo) Then
'	If (MsgBox("Save image?",vbYesNoCancel) = vbCancel) Then
'	GoTo Line1
'	End If
'	If (MsgBox("Save image?",vbYesNoCancel) = vbYes) Then
'	Set IPCurrent = IPEditor.GetIPFromName(IPEditor.ActiveIP)
'	IPCurrent.SaveImage (S & "_foto_" & Format(Now(), "dd.mm.yyyy_hh.nn.ss") & " status_OK" & ".tif")
'	Else
'	Set IPLive = IPEditor.GetIPFromName("Live")
'	IPEditor.GetIPFromName(IPEditor.ActiveIP).ActivateCamera
'	Set IPCurrent = IPEditor.GetIPFromName(IPEditor.ActiveIP)

'	IPLive.SaveImage (S & "_foto_" & Format(Now(), "dd.mm.yyyy_hh.nn.ss") & ".tif")
'	IPCurrent.SaveImage (S & "_foto_" & Format(Now(), "dd.mm.yyyy_hh.nn.ss") & " status_NG" & ".tif")
	W = MsgBox("Save image?",vbYesNoCancel)
	Select Case UCase(W)
	Case vbYes
		Set IPCurrent = IPEditor.GetIPFromName(IPEditor.ActiveIP)
		IPCurrent.SaveImage (S & "_foto_" & Format(Now(), "dd.mm.yyyy_hh.nn.ss") & " status_OK" & ".tif")
	Case vbNo
		Set IPCurrent = IPEditor.GetIPFromName(IPEditor.ActiveIP)
		IPCurrent.SaveImage (S & "_foto_" & Format(Now(), "dd.mm.yyyy_hh.nn.ss") & " status_NG" & ".tif")
	Case vbCancel
		GoTo Line1
	End Select
'	End If
	Next i
End Sub


Function Get_DirS(path As String)
Dim A() As String, D As String, U As Long
D = Dir(path, vbDirectory)
While D <> ""
  If GetAttr(path & "\" & D) And vbDirectory Then
    ReDim Preserve A(U)
    A(U) = path & D
    U = U + 1
  End If
  D = Dir
Wend
Get_DirS = A
End Function
