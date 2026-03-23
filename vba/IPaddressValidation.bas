'######### IPaddressValidation
Function IsValidIPAddress(ByVal ip As String) As Boolean
    Dim parts() As String
    Dim i As Integer
    Dim num As Integer
    
    parts = Split(ip, ".")
    
    If UBound(parts) <> 3 Then Exit Function
    
    For i = 0 To 3
        If Not IsNumeric(parts(i)) Then Exit Function
        
        num = CInt(parts(i))
        If num < 0 Or num > 255 Then Exit Function
        
        ' 先頭ゼロ防止（例: 01）
        If parts(i) <> CStr(num) Then Exit Function
    Next i
    
    IsValidIPAddress = True
End Function

Function IsValidSubnetMask(ByVal mask As String) As Boolean
    Dim parts() As String
    Dim i As Integer
    Dim num As Integer
    Dim binaryStr As String

    parts = Split(mask, ".")
    If UBound(parts) <> 3 Then Exit Function

    For i = 0 To 3
        If Not IsNumeric(parts(i)) Then Exit Function

        num = CInt(parts(i))
        If num < 0 Or num > 255 Then Exit Function

        binaryStr = binaryStr & Right("00000000" & WorksheetFunction.Dec2Bin(num), 8)
    Next i

    ' 「1が続いた後に0が続く」パターンのみ許可
    If InStr(binaryStr, "01") > 0 Then Exit Function

    ' 全部1 or 全部0は除外（必要に応じて調整）
    If binaryStr = String(32, "1") Or binaryStr = String(32, "0") Then Exit Function

    IsValidSubnetMask = True
End Function

Function IsValidNetworkAddress(ByVal ip As String, ByVal mask As String) As Boolean
    Dim ipParts() As String
    Dim maskParts() As String
    Dim i As Integer
    
    If Not IsValidIPAddress(ip) Then Exit Function
    If Not IsValidSubnetMask(mask) Then Exit Function
    
    ipParts = Split(ip, ".")
    maskParts = Split(mask, ".")
    
    For i = 0 To 3
        If (CInt(ipParts(i)) And Not CInt(maskParts(i))) <> 0 Then
            Exit Function
        End If
    Next i
    
    IsValidNetworkAddress = True
End Function
