'######### GetTimestamp
Function GetTimestamp() As String
    ' yyyy-mm-dd-HH-MM-ss 形式で現在時刻を返す
    GetTimestamp = Format(Now, "yyyy-mm-dd-HH-MM-ss")
End Function

