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
   Dim lists$(10)
  lists$(0) = "List 0"

'   lists$ = Get_DirS("C:\WWTEMP\")

   Begin Dialog UserDialog 200,120
       Text 10,10,180,15,"Please push the OK button"
       ListBox 10,25,180,60,lists$(),.list
       OKButton 80,90,40,20
   End Dialog
   Dim dlg As UserDialog
   dlg.list = 2
   Dialog dlg ' show dialog (wait for ok)
   Debug.Print dlg.list
End Sub


