Sub Main()'  saveImage()
Dim World
Dim L As String
Dim i, p As Integer
Dim IPLive As FNCImageProcess
Dim name_file As String

'Dim strFolder As String
'  L = InputBox$("Enter some text:", _
 '   "Input Box Example", "")
 '   If L = "" Then Exit Sub
  '  NString = "C:\wwtemp\" + L
'    NString = "C:\WWTEMP\"

'    L = InputBox$("Enter some text:", _
'    "Input select folders", NString)
'     ActiveWorkbook.FollowHyperlink Address:=NString, NewWindow:=True
	Shell "Explorer.exe C:\WWTEMP\", vbNormalFocus
    F$ = Dir$(L)
    NStringS = InputBox$("Enter some text:", _
    "Input selected folder of Date", F$)
	Shell "Explorer.exe C:\WWTEMP\" + NStringS, vbNormalFocus
'	D$ = NString & NStringS & "\"

	Line1:
 For i = 0 To 100
 If MsgBox("Perform image - OK, if not save - Cancel", vbOkCancel) = vbCancel Then
    Debug.Print "OK was pressed"
    Exit Sub
 End If
'	N = N + 1

' edit
	F$ = Dir$
    NStringP = InputBox$("Enter some text:", _
    "Input selected folder PCB", F$)
	For p = 1 To 100
    If (MsgBox("Save image?",vbYesNo) = vbNo) Then
'    	index = 0
    	GoTo Line1
'    ElseIf Dir$("C:\WWTEMP\" & NStringS & "\" & NStringP & "\", vbDirectory) = "" Then
'        index = index + 1
'    	Set IPLive = IPEditor.GetIPFromName("Live")
'    	IPLive.SaveImage ("C:\WWTEMP\" & NStringS & "\" & NStringP & "\" & NStringP & "_foto_" & ".tif")
	Else
		If Dir$("C:\WWTEMP\" & NStringS & "\" & NStringP & "\" & NStringP, vbDirectory) = "" Then
 '   	index = index + 1
        Set IPLive = IPEditor.GetIPFromName("Live")
        IPLive.SaveImage ("C:\WWTEMP\" & NStringS & "\" & NStringP & "\" & NStringP & "_foto_" & Format(Now(), "dd.mm.yyyy_hh.nn.ss") & ".tif")
 '   Else
 '       	index = index + 1
'			postfix = postfix + 1
 '           postfix = getNextPostfix("C:\WWTEMP\" & NStringS & "\" & NStringP & "\" & NStringP & "_foto_", True)

 '           If postfix > 0 And postfix < 13 Then

 '               Set IPLive = IPEditor.GetIPFromName("Live")
 '               name_file = "C:\WWTEMP\" & NStringS & "\" & NStringP & "\" & NStringP &  "_foto_" & postfix & ".tif"
	'			MsgBox name_file
 '               IPLive.SaveImage ("C:\WWTEMP\" & NStringS & "\" & NStringP & "\" & NStringP &  "_foto_" & postfix & ".tif")
	       	End If
        End If
 '  End If

 	Next p
 Next i
	

End Sub

'Function getNextPostfix(ByVal pathDir As String) As Integer


'Dim result As Integer, i As Integer
'For i = 1 To 15
'    If Dir$(pathDir & CStr(i)) = "" Then
'           result = i
'        Exit For
'    End If
'Next i
'getNextPostfix = result
'End Function

Function getNextPostfix(ByVal pathDir As String, Optional ByVal isFile As Boolean = False) As Integer
' ? ??????? ???????? ?????????????? ???????? isFile
Dim result As Integer, i As Integer
For i = 1 To 15
    If isFile Then ' ???? ???? isFile ??????????, ?????? ?? ???? ????
        If Dir$(pathDir & CStr(i)) = "" Then
            result = i
            Exit For
        End If
    Else ' ????? ?? ???? ???????
        If Dir$(pathDir & CStr(i), vbDirectory) = "" Then
            result = i
            Exit For
        End If
    End If
Next i
getNextPostfix = result
End Function
