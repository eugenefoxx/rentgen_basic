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
 
Sub Main()
Dim a
a = Get_DirS("C:\WWTEMP\")
MsgBox Join(a, vbLf)
End Sub
