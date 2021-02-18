'<--FNC-Version> 4 (Inserted by FNC Editor - DO NOT EDIT!)
Sub Main   '  ListBox Dialog Item Definition'
    Dim lists$()
    Dim tests$()
    Dim indays$()
    Dim IPLive As FNCImageProcess
    Dim IPCurrent As FNCImageProcess
    Dim i As Integer

'	IPEditor.GetIPFromName(IPEditor.ActiveIP).ActivateCamera

'	IPEditor.GetIPFromName(IPEditor.ActiveIP).Enable(False)   \\10.1.10.3\rentgen\Production Control\

'	lists$ = Get_DirS("C:\WWTEMP\")
	lists$ = Get_DirS("\\10.1.10.3\rentgen\Production Control\")

    Do While Fol <> ""
    	For i = LBound(lists) To UBound(lists)
        	lists$() = Dir()
		Next i
    Loop
'    lists$(0) = "List 0"
'    lists$(1) = "List 1"
'    lists$(2) = "List 2"
'    lists$(3) = "List 3"
'    lists$(4) = Fol
	Call bubbleSort(lists)
	Call ReverseArray(lists)

    Begin Dialog UserDialog 600,620   '220
        Text 10,10,18000,15,"Please push the Cancel button"
        ComboBox 10,25,18000,560,lists$(),.list$   ' 10,25,18000,160
        OKButton 10,590,40,20   '  10,190,40,20
        CancelButton 110,590,60,20  ' 110,190,60,20
    End Dialog
    Dim dlg As UserDialog
 '   dlg.list = 2
    Dialog dlg ' show dialog (wait for cancel)
    Debug.Print dlg.list$
    L = dlg.list$
 '   MsgBox L

 	tests$ = Get_DirS(L & "\")

	Do While Fol <> ""
		For i = UBound(tests) To LBound(tests)
        tests$() = Dir()
        Next i
    Loop

	Call bubbleSort(tests)

	Begin Dialog UserDialog 600,620   '220
        Text 10,10,1800,15,"Please push the button"
        ComboBox 10,25,1800,560,tests$(),.test$  ' 10,25,18000,160
        OKButton 40,590,40,20   '  10,190,40,20
        CancelButton 110,590,60,20   ' 110,190,60,20
    End Dialog
    Dim dlgS As UserDialog
 '   dlg.list = 2
  '	Line1:
    Dialog dlgS ' show dialog (wait for cancel)
'    Debug.Print dlgS.test$
    S = dlgS.test$ & "\"
'	MsgBox S

	indays$ = Get_DirS(S & "\")

	Do While Fol <> ""
		For i = UBound(indays) To LBound(indays)
        indays$() = Dir()
        Next i
    Loop

	Call bubbleSort(indays)

	Begin Dialog UserDialog 900,620   '220  900 shirina okna
        Text 10,10,1800,15,"Please push the button"
        ComboBox 10,25,1800,560,indays$(),.inday$  ' 10,25,18000,160
        OKButton 40,590,40,20   '  10,190,40,20
        CancelButton 110,590,60,20   ' 110,190,60,20
    End Dialog
    Dim dlgD As UserDialog
 '   dlg.list = 2
 	Line1:
    Dialog dlgD ' show dialog (wait for cancel)
'    Debug.Print dlgS.test$
    D = dlgD.inday$ & "\"

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
	W = MsgBox("Save image?" & vbNewLine & "Button Set: " & vbNewLine & "Yes       = status OK, " & vbNewLine & "No       = status NG, " & vbNewLine & "Cancel = Exit",vbYesNoCancel)
	Select Case UCase(W)
	Case vbYes
		Set IPCurrent = IPEditor.GetIPFromName(IPEditor.ActiveIP)
		IPCurrent.SaveImage (D & "_foto_" & Format(Now(), "dd.mm.yyyy_hh.nn.ss") & " status_OK" & ".jpg")
	Case vbNo
		Set IPCurrent = IPEditor.GetIPFromName(IPEditor.ActiveIP)
		IPCurrent.SaveImage (D & "_foto_" & Format(Now(), "dd.mm.yyyy_hh.nn.ss") & " status_NG" & ".jpg")
	Case vbCancel
		GoTo Line1
	End Select
'	End If
	Next i
End Sub


Function Get_DirS(path As String)
Dim A() As String, D As String, U As Long
D = Dir(path, vbDirectory)  '  vbDirectory
While D <> ""
  If GetAttr(path & "\" & D) And vbDirectory Then   '  vbDirectory
    ReDim Preserve A(U)
    A(U) = path & D
    U = U + 1
  End If
  D = Dir
Wend
Get_DirS = A
End Function
Function ReverseArray(vArray As Variant)
'Reverse the order of an array, so if it's already sorted
'from smallest to largest, it will now be sorted from
'largest to smallest.
Dim vTemp As Variant
Dim i As Long
Dim iUpper As Long
Dim iMidPt As Long
iUpper = UBound(vArray)
iMidPt = (UBound(vArray) - LBound(vArray)) \ 2 + LBound(vArray)
For i = LBound(vArray) To iMidPt
    vTemp = vArray(iUpper)
    vArray(iUpper) = vArray(i)
    vArray(i) = vTemp
    iUpper = iUpper - 1
Next i
End Function
Function bubbleSort(arr)
  Dim strTemp As String
  Dim i As Long
  Dim j As Long
  Dim lngMin As Long
  Dim lngMax As Long
  lngMin = LBound(arr)
  lngMax = UBound(arr)
  For i = lngMin To lngMax - 1
    For j = i + 1 To lngMax
      If arr(i) > arr(j) Then
        strTemp = arr(i)
        arr(i) = arr(j)
        arr(j) = strTemp
      End If
    Next j
  Next i
End Function
