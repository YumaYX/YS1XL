' 使用例:
'   Dim s As String: s = ReadUtf8Text("C:\Temp\file.txt")
'   Debug.Print s   ' UTF-8 のテキスト内容を文字列で取得
' 引数:
'   filePath - 読み込むファイルのフルパス
' 戻り値:
'   UTF-8 のテキスト内容を文字列で返す
'   ※ファイルが存在しない／読み取り失敗の場合は空文字
' 依存:
'   ADODB.Stream（ado moniker 経由の遅延バインド。参照設定は不要）
Function ReadUtf8Text(ByVal filePath As String) As String

    Dim stm As Object

    ReadUtf8Text = ""

    ' ファイルが存在しない場合は空文字を返す
    If Len(Dir(filePath)) = 0 Then Exit Function

    On Error GoTo CleanUp

    Set stm = CreateObject("ADODB.Stream")

    With stm
        .Type = 2
        .Charset = "UTF-8"
        .Open
        .LoadFromFile filePath
        ReadUtf8Text = .ReadText
        .Close
    End With

CleanUp:
    ' エラー時もストリームを確実に解放する
    On Error Resume Next
    If Not stm Is Nothing Then stm.Close
    Set stm = Nothing
    On Error GoTo 0

End Function
