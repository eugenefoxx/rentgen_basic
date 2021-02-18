'<--FNC-Version> 4 (Inserted by FNC Editor - DO NOT EDIT!)

Sub Main

Dim lists$()

lists$ = Get_DirS("C:\WWTEMP\")


Do While Fol <> ""
        lists$() = Dir()
    Loop


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
End Sub


Function Get_DirS(path As String)
Dim A() As String, D As String, U As Long
D = Dir(path, vbDirectory)
While D <> ""
  If GetAttr(path & "\" & D) And vbDirectory Then
    ReDim Preserve A(U)
    A(U) = path & D
    U = U + 1
  End If
  D = Dir
Wend
Get_DirS = A
End Function
