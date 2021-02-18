Sub Main
    F$ = Dir$("C:\WWTEMP\*", vbDirectory)
    While F$ <> ""
        Debug.Print F$
        F$ = Dir$()
    Wend
MsgBox F
End Sub
