Sub Main
    Dim combos$(100000)
	Dim L

	F$ = "C:\WWTEMP\14.03.2019. - 15.58"

    combos$(0) = "Combo 0"
    combos$(1) = "Combo 1"
    combos$(2) = "Combo 2"
    combos$(3) = "Combo 3"
	combos$(4) = F	
    Begin Dialog UserDialog 200,300   '120
        Text 10,10,180,15,"Please push the OK button"
'		ComboBox 10,25,180,60,F$(),.F$
        ComboBox 10,25,180,60,combos$(),.combo$
        OKButton 80,90,40,20
    End Dialog
    Dim dlg As UserDialog
    dlg.combo$ = "none"
    Dialog dlg ' show dialog (wait for ok)
    Debug.Print dlg.combo$
	L = dlg.combo$
	MsgBox L
	Shell "Explorer.exe " + L, vbNormalFocus

End Sub
