Dim lists$()

Sub Main
    ReDim lists$(0)
'    lists$(0) = "List 0"
'    lists$ = Get_DirS("C:\WWTEMP\")
    lists$ = Get_DirS("\\10.1.10.3\rentgen\Production Control\")

	Do While Fol <> ""
        lists$() = Dir()
    Loop

    Begin Dialog UserDialog 900,620,.DialogFunc
        Text 10,10,1800,15,"Please push the OK button"
        ListBox 10,25,1800,560,lists(),.list
        OKButton 40,590,40,20
'        PushButton 110,590,60,20,"&Change"
    End Dialog
    Dim dlg As UserDialog
'    dlg.list = 2
    Dialog dlg ' show dialog (wait for ok)
    Debug.Print dlg.list
End Sub

Function DialogFunc%(DlgItem$, Action%, SuppValue%)
    Select Case Action%
    Case 2 ' Value changing or button pressed
        If DlgItem$ = "Change" Then
            Dim N As Integer
            N = UBound(lists$)+1
            ReDim Preserve lists$(N)
            lists$(N) = "List " & N
            DlgListBoxArray "list",lists$()
            DialogFunc% = True 'do not exit the dialog
        End If
    End Select
End Function
Function Get_DirS(path As String)
Dim A() As String, D As String, U As Long
D = Dir(path, vbDirectory)  '  vbDirectory
While D <> ""
  If GetAttr(path & "\" & D) And vbDirectory Then   '  vbDirectory
    ReDim Preserve A(U)
    A(U) = path & D
    U = U + 1
  End If
  D = Dir
Wend
Get_DirS = A
End Function
