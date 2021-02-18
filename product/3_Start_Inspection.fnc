'<--FNC-Version> 4 (Inserted by FNC Editor - DO NOT EDIT!)

Sub Main()
Dim Insert, SAP
Dim S

Manipulator.OpenDoor(True)
Line1:
SAPFOLDER = InputBox$("Scanning ID:", _
        "Input Box ID", "")
        If SAPFOLDER = "" Then Exit Sub
        If Mid(SAPFOLDER, 1, 1) = "P" Then
       '   MsgBox "OK"

'          NString = "C:\WWTEMP\SAP\" + Mid(SAPFOLDER, 2, 7)
		NString = "\\10.1.10.3\rentgen\Components\" + Mid(SAPFOLDER, 2, 7)
            If Dir$(NString, vbDirectory) = "" Then
               MkDir NString
            End If
        Else
          MsgBox "Barcode not scanned correctly. Please again."
          GoTo Line1
        End If
 '    D = Mid(SAPFOLDER, 2, 7)

' Insert = "P1017036LK339470161R1000201639Q02000D20190211"
' Insert = "P00000000000000000000000000000000000000000000"

' NString = "C:\WWTEMP\SAP\" + Mid(SAPFOLDER, 2, 7)

'    If Dir$(NString, vbDirectory) = "" Then
'        MkDir NString
'    End If


'    NStringS = NString & "\" & NumStr2 & "_" & "\" & S
 '   Ss = Mid(S8, 16)
    SubSAP = NString & "\" & "Lot " & Mid(SAPFOLDER, 10, 10) & " ID " & Mid(SAPFOLDER, 21, 10) & Format(Now(), " dd.mm.yyyy. - hh.nn.ss")
	MkDir SubSAP
'		Manipulator.OpenDoor(True)

   		PhotosubSAP = SubSAP & "\"

   		For i = 0 To 100
 '  		Manipulator.OpenDoor(True)
'   		W = MsgBox("Save image?" & vbNewLine & "Button Set: " & vbNewLine & "Yes       = status OK, " & vbNewLine & "No       = status NG, " & vbNewLine & "Cancel = Exit",vbOkCancel) 'vbYesNoCancel
	W = MsgBox("Save image?",vbOkCancel)
	Select Case UCase(W)
'	Case vbYes
'		Set IPCurrent = IPEditor.GetIPFromName(IPEditor.ActiveIP)
'		IPCurrent.SaveImage (PhotosubSAP & "_foto_" & Format(Now(), "dd.mm.yyyy_hh.nn.ss") & " status_OK" & ".tif")
'	Case vbNo
'		Set IPCurrent = IPEditor.GetIPFromName(IPEditor.ActiveIP)
'		IPCurrent.SaveImage (PhotosubSAP & "_foto_" & Format(Now(), "dd.mm.yyyy_hh.nn.ss") & " status_NG" & ".tif")
'	Case vbCancel
'		GoTo Line1
'	End Select
'	End If
	Case vbOK
		Set IPCurrent = IPEditor.GetIPFromName(IPEditor.ActiveIP)
		IPCurrent.SaveImage (PhotosubSAP & "_foto_" & Format(Now(), "dd.mm.yyyy_hh.nn.ss") & ".jpg")
	Case vbCancel
		Manipulator.OpenDoor(True)
		GoTo Line1
	End Select

	Next i

' Insert = NString

'SAP = Mid(SAPFOLDER, 2, 7)
'MsgBox (SAP)

End Sub
