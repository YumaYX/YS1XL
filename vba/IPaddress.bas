' Debug.Print IsValidIPAddress("192.168.0.1")       ' True
' Debug.Print IsValidIPAddress("192.168.0.999")     ' False
' 引数:
'   ip - 検証するIPアドレス
' 戻り値:
'   有効なら True / 無効なら False
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

' Debug.Print IsValidSubnetMask("255.255.255.0")    ' True
' Debug.Print IsValidSubnetMask("255.0.255.0")      ' False（1の連続でない）
' 引数:
'   mask - 検証するサブネットマスク
' 戻り値:
'   有効なら True / 無効なら False
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

' Debug.Print IsValidNetworkAddress("192.168.0.1", "255.255.255.0")   ' True
' Debug.Print IsValidNetworkAddress("192.168.1.1", "255.255.255.0")   ' False
' 引数:
'   ip   - 検証するIPアドレス
'   mask - サブネットマスク
' 戻り値:
'   有効なら True / 無効なら False
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

' Debug.Print CIDR2Mask(24)   ' 255.255.255.0
' Debug.Print CIDR2Mask(8)    ' 255.0.0.0
' 引数:
'   cidr - プレフィックス長（0～32）
' 戻り値:
'   ドット区切りのサブネットマスク文字列
Function CIDR2Mask(cidr As Integer) As String
    Dim i As Integer
    Dim mask(3) As Integer
    Dim bits As Integer

    bits = cidr

    For i = 0 To 3
        If bits >= 8 Then
            mask(i) = 255
            bits = bits - 8
        ElseIf bits > 0 Then
            mask(i) = 256 - 2 ^ (8 - bits)
            bits = 0
        Else
            mask(i) = 0
        End If
    Next i

    CIDR2Mask = mask(0) & "." & mask(1) & "." & mask(2) & "." & mask(3)
End Function

' Debug.Print Mask2CIDR("255.255.255.0")   ' 24
' Debug.Print Mask2CIDR("255.0.0.0")       ' 8
' 引数:
'   mask - ドット区切りのサブネットマスク
' 戻り値:
'   プレフィックス長（0～32）
Function Mask2CIDR(mask As String) As Integer
    Dim parts() As String
    Dim i As Integer
    Dim val As Integer
    Dim cidr As Integer

    parts = Split(mask, ".")

    For i = 0 To 3
        val = CInt(parts(i))

        Do While val > 0
            cidr = cidr + (val And 1)
            val = val \ 2
        Loop
    Next i

    Mask2CIDR = cidr
End Function




