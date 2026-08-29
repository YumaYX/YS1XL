' Dim ts As String: ts = GetTimestamp()
' Debug.Print ts   ' 2026-08-19-09-30-00
Function GetTimestamp() As String
    ' yyyy-mm-dd-HH-MM-ss 形式で現在時刻を返す
    GetTimestamp = Format(Now, "yyyy-mm-dd-HH-MM-ss")
End Function





