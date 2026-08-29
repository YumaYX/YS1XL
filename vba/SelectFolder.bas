' 使用例:
'   Dim path As String: path = SelectFolder()
'   If path = "" Then Exit Sub     ' キャンセル時は空文字
' 戻り値:
'   選択したフォルダのフルパス（キャンセル時は ""）
Function SelectFolder() As String
    Dim fd As Object
    Set fd = Application.FileDialog(4)
    With fd
        .Title = "フォルダを選択してください"
        .AllowMultiSelect = False

        If .Show = -1 Then
            SelectFolder = .SelectedItems(1)
        Else
            SelectFolder = ""
        End If
    End With
    Set fd = Nothing
End Function

