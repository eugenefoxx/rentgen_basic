Sub Main
Dim FNames As String
Dim arr()
Dim i As Long, j As Long
Dim lists$()

' ReDim arr(10000) 'set greater than possible number of files
lists$ = Get_DirS("\\10.1.10.3\rentgen\Production Control\")

Do While Fol <> ""
'	arr(i) = lists$
'	i = i + 1
	For i = LBound(lists) To UBound(lists)
'	lists(i) = i
    lists$() = Dir()
    Next i
Loop

Call ReverseArray(lists)
'ReDim Preserve arr(i)
'QuickSort arr(), LBound(arr), UBound(arr)  ' QuickSort arr(), LBound(arr), UBound(arr)
'For j = 0 To i
 ' Cells(j + 1, 2) = arr(j)
' lists$() = arr(j)
'MsgBox arr(j)
 
'Next

Begin Dialog UserDialog 600,620   '220
        Text 10,10,18000,15,"Please push the Cancel button"
        ComboBox 10,25,18000,560,lists$(), .list   ' 10,25,18000,160
        OKButton 10,590,40,20   '  10,190,40,20
        CancelButton 110,590,60,20  ' 110,190,60,20
    End Dialog
    Dim dlg As UserDialog
 '   dlg.list = 2
    Dialog dlg ' show dialog (wait for cancel)
    Debug.Print dlg.list




End Sub


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

Private Sub QuickSort(strArray(), intBottom As Integer, intTop As Integer)
Dim strPivot As String, strTemp As String
Dim intBottomTemp As Integer, intTopTemp As Integer
intBottomTemp = intBottom
intTopTemp = intTop
strPivot = strArray((intBottom + intTop) \ 2)
While (intBottomTemp <= intTopTemp)
While (strArray(intBottomTemp) < strPivot And intBottomTemp < intTop)
intBottomTemp = intBottomTemp + 1
Wend
While (strPivot < strArray(intTopTemp) And intTopTemp > intBottom)
intTopTemp = intTopTemp - 1
Wend
If intBottomTemp < intTopTemp Then
strTemp = strArray(intBottomTemp)
strArray(intBottomTemp) = strArray(intTopTemp)
strArray(intTopTemp) = strTemp
End If
If intBottomTemp <= intTopTemp Then
intBottomTemp = intBottomTemp + 1
intTopTemp = intTopTemp - 1
End If
Wend
'the function calls itself until everything is in good order
If (intBottom < intTopTemp) Then QuickSort strArray, intBottom, intTopTemp
If (intBottomTemp < intTop) Then QuickSort strArray, intBottomTemp, intTop
End Sub
