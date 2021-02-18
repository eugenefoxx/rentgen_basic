'<--FNC-Version> 4 (Inserted by FNC Editor - DO NOT EDIT!)


Sub Multiscan()

Dim L, S
Dim g, i As Integer
Dim postfix As Integer

For g = 0 To 100
'    For i = 0 To 100

If MsgBox("If we continue scanning - OK, if we complete - Cancel", vbOkCancel) = vbCancel Then
    Debug.Print "OK was pressed"
    Exit Sub
End If
    L = InputBox$("Enter some text:", _
    "Input Box Example", "", g)
    If L = "" Then Exit Sub
    NString = "C:\WWTEMP\" + L

    If Dir$(NString, vbDirectory) = "" Then
        MkDir NString
    End If

    For i = 0 To 100
    If MsgBox("If we Continue scanning - OK, if we Complete - Cancel", vbOkCancel) = vbCancel Then
        Debug.Print "OK was pressed"
        Exit Sub
    End If
        S = InputBox$("Enter some text:", _
        "Input Box Example", "", i)
        NStringS = NString & "\" + S

        If Dir$(NStringS, vbDirectory) = "" Then
        MkDir NStringS
        Else
            If Dir$(NStringS & "_repair1", vbDirectory) = "" Then
                MkDir NStringS & "_repair1"
            Else
                postfix = getNextPostfixM(NStringS & "_repair")
                If postfix > 1 And postfix < 6 Then
                    MkDir (NStringS & "_repair" & postfix)
                End If
            End If
        Debug.Print S

    End If

   Next i
Next g
End Sub

Function getNextPostfixM(ByVal pathDir As String) As Integer


Dim result As Integer, i As Integer
For i = 1 To 5
    If Dir$(pathDir & CStr(i), vbDirectory) = "" Then
           result = i
        Exit For
    End If
Next i
getNextPostfixM = result
End Function

Sub Main()
Dim L
Dim g As Integer
Dim postfix As Integer

For g = 0 To 100

If MsgBox("If we continue scanning - OK, if we complete - Cancel", vbOkCancel) = vbCancel Then
    Debug.Print "Cancel was pressed"
    Exit Sub
End If
    L = InputBox$("Enter some text:", _
    "Input Box Example", "", g)
    If L = "" Then Exit Sub
    NString = "C:\WWTEMP\" + L

    If Dir$(NString, vbDirectory) = "" Then
        MkDir NString
    Else
        If Dir$(NString & "_repair1", vbDirectory) = "" Then
            MkDir NString & "_repair1"
        Else
            postfix = getNextPostfix(NString & "_repair")
            If postfix > 1 And postfix < 6 Then
                MkDir (NString & "_repair" & postfix)
            End If
        End If
    Debug.Print L
End If
Next
End Sub

Function getNextPostfix(ByVal pathDir As String) As Integer
Dim result As Integer, i As Integer
For i = 1 To 5
    If Dir$(pathDir & CStr(i), vbDirectory) = "" Then
        result = i
        Exit For
    End If
Next i
getNextPostfix = result
End Function





    ' 705
'	FNCMsgBox 944,705,L$
'	NString ="C:\WWTEMP\"+L

'	MkDir NString
'	If NString <> "" Then

'    If Dir$(NString, vbDirectory) = "" Then  ' If Dir$(NString, vbDirectory) <> "" Then
'		MkDir NString

'	MsgBox "The Folder already exists!"
'	ElseIf Dir$(NString, vbDirectory)= "" Then
'		MkDir NString & "_repair1"
'	Else
	'If NString <> "" Then
'		MkDir NString & "_repair2"
'	MsgBox "The Folder already exists!"
'	MsgBox "The Folder not exists!"
'	MkDir NString
'	Exit Sub
'	End If
'	Exit Sub
'	End If
'	Exit Sub
'	Next
'	Exit Sub
'	Exit Sub
'
'	MsgBox "Please press OK button"
'    If MsgBox("Please press OK button",vbOkCancel) = vbOK Then
'        Debug.Print "OK was pressed"
'    Else
'    Debug.Print "Cancel was pressed"
'    End If
'		End Sub

