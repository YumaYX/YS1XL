' 引数:
'   folderPath - 対象フォルダパス（末尾の \ はなくても自動補完）
'   pattern    - ファイル名のワイルドカードパターン（例: *.txt）
'
' 戻り値:
'   パターンに一致したファイルのフルパスを格納した Collection
'
' 使用例:
'   Dim col As Collection: Set col = GetFiles("C:\Temp", "*.txt")
'   For Each f In col: Debug.Print f: Next
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


' 引数:
'   folderPath - 対象フォルダパス（末尾の \ はなくても自動補完）
'   pattern    - ファイル名のワイルドカードパターン（例: *.txt）
'
' 戻り値:
'   サブフォルダを含めてパターンに一致したファイルのフルパスを格納した Collection
'   （対象フォルダが存在しない場合は空の Collection）
'
' 使用例:
'   Dim col As Collection: Set col = GetFilesRecursive("C:\Temp", "*.txt")
'   For Each f In col: Debug.Print f: Next
Function GetFilesRecursive(folderPath As String, pattern As String) As Collection
    Dim col As New Collection
    Dim fso As Object

    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FolderExists(folderPath) Then
        GetFilesRecursive_Add fso.GetFolder(folderPath), pattern, col
    End If

    Set GetFilesRecursive = col
End Function


' 引数:
'   folder  - 走査対象の FSO Folder オブジェクト
'   pattern - ファイル名のワイルドカードパターン（例: *.txt）
'   col     - 一致したファイルのフルパスを追加する Collection（ByRef）
'
' 戻り値: なし（col に直接追加される）
Private Sub GetFilesRecursive_Add( _
    ByVal folder As Object, _
    ByVal pattern As String, _
    ByRef col As Collection)

    Dim file As Object
    Dim subFolder As Object

    ' ファイル
    For Each file In folder.Files
        If file.Name Like pattern Then
            col.Add file.Path
        End If
    Next file

    ' サブフォルダ
    For Each subFolder In folder.SubFolders
        GetFilesRecursive_Add subFolder, pattern, col
    Next subFolder

End Sub




