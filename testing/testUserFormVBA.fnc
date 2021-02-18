Private Sub Main()
'Filename = Dir("C:\Users\???????\Desktop" & "\*.xlsx", vbNormal)
Filename = Dir("C:\Users\???????\Desktop\" & "\", vbDirectory)
Do While Len(Filename) > 0

'UserForm1.ListBox1.AddItem Filename
UserForm1.ComboBox1.AddItem Filename
Filename = Dir()

Loop
F$ = CommandButton1_Click()
FilenameS = Dir("C:\Users\???????\Desktop\" & "\" & Filename, vbDirectory)
Do While Len(FilenameS) > 0

'UserForm1.ListBox1.AddItem Filename
UserForm1.ComboBox2.AddItem FilenameS
FilenameS = Dir()

Loop
End Sub

Private Sub CommandButton1_Click()
If UserForm1.ComboBox1.ListIndex < 0 Then
MsgBox ("SELECT A FILE!!!!..... Jeez")
Else

Filename = UserForm1.ComboBox1.list(UserForm1.ComboBox1.ListIndex)
Filename = "C:\Users\???????\Desktop\" & Filename
' Workbooks.Open (Filename)
'Shell "Explorer.exe C:\Users\???????\Desktop\" & Filename, vbNormalFocus
FilenameS = Filename
End If
End Sub
