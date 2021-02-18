Sub Main
    F$ = Dir$("*.*")
    While F$ <> ""
        Debug.Print F$;" ";GetAttr(F$)
        F$ = Dir$()
    Wend
End Sub