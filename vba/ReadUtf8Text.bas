'######### ReadUtf8Text
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

