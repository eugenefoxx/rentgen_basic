Sub Main()
    Begin Dialog UserDialog 200,120
        Text 10,10,180,15,"Please push the OK button"
        OKButton 80,90,40,20
    End Dialog
    Dim dlg As UserDialog
    Dialog dlg ' show dialog (wait for ok)
End Sub