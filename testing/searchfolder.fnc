Option Compare Text

Private iCollection As Collection

Private Sub Test()
    Dim iFileObject As Object, iFolder As Object
    
    iStartPath$ = "C:\"
    iSearchPath$ = "Common FileS"
    
    Set iFileObject = CreateObject("Scripting.FileSystemObject")
    Set iFolder = iFileObject.GetFolder(iStartPath$)
    Set iCollection = New Collection
    
    FindSubFolder iFolder, iSearchPath$
    
    If iCollection.Count = 0 Then
       MsgBox "Folder not found", , ""
    Else
       For Each iFolder In iCollection
           MsgBox iFolder.Path, , ""
       Next
    End If
End Sub

Private Sub FindSubFolder(iFolder As Object, iSearchPath$)
     For Each iFolder In iFolder.SubFolders
         If iFolder.Name = iSearchPath Then
            iCollection.Add iFolder ' .Path
         Else
            FindSubFolder iFolder, iSearchPath
         End If
     Next
End Sub
