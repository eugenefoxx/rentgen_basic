'<--FNC-Version> 4 (Inserted by FNC Editor - DO NOT EDIT!)

Sub Main()
Dim L, S
Dim g, i As Integer
Dim postfix As Integer


' Line1:
For g = 0 To 100
'If MsgBox("Create the Folder with Date. If we continue scanning Pattern - OK, if we complete Pattern - Cancel", vbOkCancel) = vbCancel Then
'    Debug.Print "OK was pressed"
'	MsgBox "Close Door"
'    Exit Sub
' End If

'	MsgBox "Open Door"
	Manipulator.OpenDoor(True)

'    L = InputBox$("Enter some text:", _
'    "Input Box Example", "", g)
'    If L = "" Then Exit Sub '
'    N = N + 1
'    NString = "C:\wwtemp\" + L  CStr(Date)

'     NString = "C:\WWTEMP\" & Format(Now(), "dd.mm.yyyy. - hh.nn")  \\10.1.10.3\rentgen\
'    NString = "C:\WWTEMP\" & Format(Now(), "yyyy.mm.dd.- hh.nn")
	NStringDay = "\\10.1.10.3\rentgen\Production Control\" & Format(Now(), "yyyy.mm.dd.")
            If Dir$(NStringDay, vbDirectory) = "" Then ' ?????????, ???? ?? ???, ?? ??????? ?????
               MkDir NStringDay
            End If

'	NString = "\\10.1.10.3\rentgen\Production Control\" & Format(Now(), "yyyy.mm.dd.- hh.nn")
	NString = NStringDay & "\" & Format(Now(), "yyyy.mm.dd.- hh.nn")
	Check = NString & "\" & "Check"

    If Dir$(NString, vbDirectory) = "" Then
        MkDir NString
		MkDir Check
		SetAttr Check,vbHidden
    End If

	Line1:
    If MsgBox("If we Continue scanning PCB - OK, if we Complete PCB - Cancel", vbOkCancel) = vbCancel Then
        Debug.Print "OK was pressed"
 '       Manipulator.CloseDoor(True)
 '       GoTo Line1
        Exit Sub
	End If
	For i = 0 To 100
		Line2:
        S = InputBox$("Scan form:", _
        "Please scan the PCB", "", i)
 '       If S = "" Then Exit Sub
 		If S = "" Then GoTo Line1
 '		Do While N < 100

		Checks = Check & "\" & S
		If Dir$(Checks, vbDirectory) = "" Then
			MkDir Checks
		Else
			MsgBox "ERROR: DOUBLE!!!"
			GoTo Line2
		End If

        N = N + 1
        NumStr = CStr(Format(N,"000"))

        NStringS = NString & "\" & NumStr & "_" & S
	 '	Check = Right(NStringS,16)

            If Dir$(NStringS, vbDirectory) = "" Then
 '			If Dir$(NStringS, vbDirectory) <> Right(Dir$(NStringS, vbDirectory),16) Then
 '					If Right(NStringS, 16) <> S Then
 '					If Check = "" Then
                 MkDir NStringS
	'			MsgBox "ERROR.Set EN (English) layout."
            Else
	'			MkDir NStringS
	 			MsgBox "ERROR.Set EN (English) layout."
            End If

	'	Loop
    Next i
 '   GoTo Line1
Next g

End Sub 
