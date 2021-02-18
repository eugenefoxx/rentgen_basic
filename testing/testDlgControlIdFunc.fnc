Sub Main
    Begin Dialog UserDialog 200,120,.DialogFunc
        Text 10,10,180,15,"Please push the OK button"
        TextBox 10,40,180,15,.Text
        OKButton 30,90,60,20
        PushButton 110,90,60,20,"&Hello"
    End Dialog
    Dim dlg As UserDialog
    Debug.Print Dialog(dlg)
End Sub

Function DialogFunc%(DlgItem$, Action%, SuppValue%)
    Debug.Print "Action=";Action%
    Select Case Action%
    Case 1 ' Dialog box initialization
        Beep
    Case 2 ' Value changing or button pressed
        If DlgItem$ = "Hello" Then
            DialogFunc% = True 'do not exit the dialog
        End If
    Case 4 ' Focus changed
        Debug.Print "DlgFocus=""";DlgFocus();""""
        Debug.Print "DlgControlId(";DlgItem$;")=";
        Debug.Print DlgControlId(DlgItem$)
    End Select
End Function