Sub Main
    Dim lists$(3)
    lists$(0) = "List 0"
    lists$(1) = "List 1"
    lists$(2) = "List 2"
    lists$(3) = "List 3"
    Begin Dialog UserDialog 200,120
        Text 10,10,180,15,"Please push the OK button"
        DropListBox 10,25,180,60,lists$(),.list1
        DropListBox 10,50,180,60,lists$(),.list2,1
        OKButton 80,90,40,20
    End Dialog
    Dim dlg As UserDialog
    dlg.list1 = 2       ' list1 is a numeric field
    dlg.list2 = "xxx"   ' list2 is a string field
    Dialog dlg ' show dialog (wait for ok)
    Debug.Print lists$(dlg.list1)
    Debug.Print dlg.list2
End Sub