Option Explicit
 
Function Get_DirS(path As String)
Dim a() As String, D As String, U As Long
D = Dir(path, vbDirectory)
While D <> ""
  If GetAttr(path & "\" & D) And vbDirectory Then
    ReDim Preserve a(U)
    a(U) = path & D
    U = U + 1
  End If
  D = Dir
Wend
Get_DirS = a
End Function
 
'Sub Get_DirS_Example()
'Dim a
'a = Get_DirS("C:\WWTEMP\")


' MsgBox Join(a, vbLf)
'End Sub



Sub Main
'    Dim list$(100000)
    Dim a, b
'    Dim List, i As Integer
'    NString = "C:\WWTEMP\" + L
'    For List i = 0 To 1000
'        L = i + 1
'    Next List
'    list = NString
    a = Get_DirS("C:\WWTEMP\")
'    b = Join(a, vbLf)
	MsgBox Join(a, vbLf)
'    list = b
'    Begin Dialog UserDialog 200,120
'    Text 10,10,180,15,"Please push the OK button"
'    ListBox 10,25,180,60, list&(),.list
'    OKButton 80,90,40,20
'End Dialog
'Dim dlg As UserDialog
'Dialog dlg 'show dialog (wait for ok)
'Debug.Print dlg.list

End Sub


