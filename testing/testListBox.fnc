Sub Main
    Dim lists$(100000)
	Fol$ = "C:\WWTEMP\14.03.2019. - 15.58"

    lists$(0) = "List 0"
    lists$(1) = "List 1"
    lists$(2) = "List 2"
    lists$(3) = "List 3"
    lists$(4) = Fol
    Begin Dialog UserDialog 700,120
        Text 10,10,700,15,"Please push the OK button"
        ListBox 10,25,180,10000,lists$(),.list
        OKButton 80,90,40,20
    End Dialog
    Dim dlg As UserDialog
'    dlg.list = 2
    Dialog dlg ' show dialog (wait for ok)
    Debug.Print dlg.list
End Sub
