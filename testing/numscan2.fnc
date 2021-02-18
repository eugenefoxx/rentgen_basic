Sub Main()
Dim L, S
Dim g, i As Integer
Dim postfix As Integer


' Line1:
For g = 0 To 100
If MsgBox("Create the Folder with Date. If we continue scanning Pattern - OK, if we complete Pattern - Cancel", vbOkCancel) = vbCancel Then
    Debug.Print "OK was pressed"
    Exit Sub

End If
'    L = InputBox$("Enter some text:", _
'    "Input Box Example", "", g)
'    If L = "" Then Exit Sub '
'    N = N + 1
'    NString = "C:\wwtemp\" + L  CStr(Date)

     NString = "C:\WWTEMP\" & Format(Now(), "dd.mm.yyyy. - hh.nn")
    If Dir$(NString, vbDirectory) = "" Then
        MkDir NString
    End If

	Line1:
    If MsgBox("If we Continue scanning PCB - OK, if we Complete PCB - Cancel", vbOkCancel) = vbCancel Then
        Debug.Print "OK was pressed"
 '       GoTo Line1
        Exit Sub
	End If
	For i = 0 To 100
        S = InputBox$("Scan form:", _
        "Please scan the PCB", "", i)
 '       If S = "" Then Exit Sub
 		If S = "" Then GoTo Line1
        N = N + 1
        NStringS = NString & "\" & N & "_" & S


            If Dir$(NStringS, vbDirectory) = "" Then
                MkDir NStringS
            End If

    Next i
 '   GoTo Line1
Next g
End Sub
