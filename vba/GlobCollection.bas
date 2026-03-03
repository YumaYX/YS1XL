'######### GlobCollection
Function GlobCollection(folderPath As String, pattern As String) As Collection
    Dim col As New Collection
    Dim fileName As String: fileName = Dir(folderPath & pattern)
    
    Do While fileName <> ""
        col.Add folderPath & fileName
        fileName = Dir()
    Loop
    Set GlobCollection = col
End Function

