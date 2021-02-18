Sub Main
    Begin Dialog UserDialog 200,120
        Text 10,10,180,30,"Please push the Cancel button"
        OKButton 40,90,40,20
        CancelButton 110,90,60,20
    End Dialog
    Dim dlg As UserDialog
    Dialog dlg ' show dialog (wait for cancel)
    Debug.Print "Cancel was not pressed"
End Sub