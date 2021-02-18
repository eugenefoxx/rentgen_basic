
Sub Main

Dim FNames As String
Dim arr()
Dim i As Long, j As Long
Dim lists$()

ReDim arr(10000) 'set greater than possible number of files
lists$ = Get_DirS("\\10.1.10.3\rentgen\Production Control\")




End Sub


Function ReverseArray(vArray As Variant)
'Reverse the order of an array, so if it's already sorted
'from smallest to largest, it will now be sorted from
'largest to smallest.
Dim vTemp As Variant
Dim i As Long
Dim iUpper As Long
Dim iMidPt As Long
iUpper = UBound(vArray)
iMidPt = (UBound(vArray) - LBound(vArray)) \ 2 + LBound(vArray)
For i = LBound(vArray) To iMidPt
    vTemp = vArray(iUpper)
    vArray(iUpper) = vArray(i)
    vArray(i) = vTemp
    iUpper = iUpper - 1
Next i
End Function 

Sub Reverse_Example2()
Dim v(-5 To 5) As Variant
For i = LBound(v) To UBound(v)
    v(i) = i
Next i
Call ReverseArray(v)
'From here on, the array "v" is in reverse order (5,4,3...-4,-5)
End Sub



Sub Main
    Dim lists$(3)
    lists$(0) = "List0"
    lists$(1) = "List1"
    lists$(2) = "List2"
    lists$(3) = "List3"

	For i = LBound(lists) To UBound(lists)
	MsgBox lists(i)
	lists(i) = i
	Next i

	Call ReverseArray(lists)

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

Function ReverseArray(vArray As Variant)
'Reverse the order of an array, so if it's already sorted
'from smallest to largest, it will now be sorted from
'largest to smallest.
Dim vTemp As Variant
Dim i As Long
Dim iUpper As Long
Dim iMidPt As Long
iUpper = UBound(vArray)
iMidPt = (UBound(vArray) - LBound(vArray)) \ 2 + LBound(vArray)
For i = LBound(vArray) To iMidPt
    vTemp = vArray(iUpper)
    vArray(iUpper) = vArray(i)
    vArray(i) = vTemp
    iUpper = iUpper - 1
Next i
End Function