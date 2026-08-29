' Dim col As Collection: Set col = GetFiles("C:\Temp", "*.txt")
' For Each f In col: Debug.Print f: Next
Function GetFiles(folderPath As String, pattern As String) As Collection
    Dim col As New Collection
    Dim fileName As String

    If Right(folderPath, 1) <> "\" Then
        folderPath = folderPath & "\"
    End If

    fileName = Dir(folderPath & pattern)

    Do While fileName <> ""
        col.Add folderPath & fileName
        fileName = Dir()
    Loop

    Set GetFiles = col
End Function


' GetFilesRecursive: サブフォルダも含めて検索する
Function GetFilesRecursive(folderPath As String, pattern As String) As Collection
    Dim col As New Collection

    If Right(folderPath, 1) <> "\" Then
        folderPath = folderPath & "\"
    End If

    GetFilesRecursive_Add col, folderPath, pattern

    Set GetFilesRecursive = col
End Function

Private Sub GetFilesRecursive_Add(col As Collection, folderPath As String, pattern As String)
    Dim fileName As String
    Dim subFolder As String

    fileName = Dir(folderPath & pattern)
    Do While fileName <> ""
        col.Add folderPath & fileName
        fileName = Dir()
    Loop

    subFolder = Dir(folderPath & "*.*", vbDirectory)
    Do While subFolder <> ""
        If subFolder <> "." And subFolder <> ".." Then
            If (GetAttr(folderPath & subFolder) And vbDirectory) = vbDirectory Then
                GetFilesRecursive_Add col, folderPath & subFolder & "\", pattern
            End If
        End If
        subFolder = Dir()
    Loop
End Sub




