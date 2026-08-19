'######### ReadUtf8Text
' Dim s As String: s = ReadUtf8Text("C:\Temp\file.txt")
' Debug.Print s   ' UTF-8 のテキスト内容を文字列で取得
Function ReadUtf8Text(filePath As String) As String
    
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    
    With stm
        .Type = 2
        .Charset = "UTF-8"
        .Open
        .LoadFromFile filePath
        ReadUtf8Text = .ReadText
        .Close
    End With
    
    Set stm = Nothing

End Function


