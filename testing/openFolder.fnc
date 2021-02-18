Sub Main()  ' openFolder()
Dim L As String
'Dim strFolder As String

  L = InputBox$("Enter some text:", _
    "Input Box Example", "")
    If L = "" Then Exit Sub
    NString = "C:\wwtemp\" + L



 Shell "Explorer.exe C:\WWTEMP\" + L, vbNormalFocus
'strFolder = ???C:\exceltrainingvideos\???
' ActiveWorkbook.FollowHyperlink Address:=NString, NewWindow:=True
' Shell ???Explorer.exe C:\wwtemp\???, vbNormalFocus
 '   Shell (???Explorer.exe""C:\wwtemp\", vbNormalFocus)
End Sub
