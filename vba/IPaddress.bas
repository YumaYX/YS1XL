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
'   全0（0.0.0.0 = /0）と全1（255.255.255.255 = /32）も有効として扱う
'   （CIDR2Mask(0) / CIDR2Mask(32) と整合させるため）
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

        binaryStr = binaryStr & Dec2Bin8(num)
    Next i

    ' 「1が続いた後に0が続く」パターンのみ許可する
    If InStr(binaryStr, "01") > 0 Then Exit Function

    IsValidSubnetMask = True
End Function

' 0〜255 を 8桁の 2進文字列に変換する
' WorksheetFunction.Dec2Bin に依存せず自身の計算のみで求める
Private Function Dec2Bin8(ByVal num As Integer) As String
    Dim result As String
    Dim i As Integer

    For i = 7 To 0 Step -1
        If (num And 2 ^ i) > 0 Then
            result = result & "1"
        Else
            result = result & "0"
        End If
    Next i

    Dec2Bin8 = result
End Function

' Debug.Print IsValidNetworkAddress("192.168.0.0", "255.255.255.0")   ' True
' Debug.Print IsValidNetworkAddress("192.168.0.1", "255.255.255.0")   ' False
' 引数:
'   ip   - 検証するIPアドレス
'   mask - サブネットマスク
' 戻り値:
'   ネットワークアドレスなら True / そうでなければ False
Function IsValidNetworkAddress(ByVal ip As String, ByVal mask As String) As Boolean
    Dim ipParts() As String
    Dim maskParts() As String
    Dim i As Integer

    If Not IsValidIPAddress(ip) Then Exit Function
    If Not IsValidSubnetMask(mask) Then Exit Function

    ipParts = Split(ip, ".")
    maskParts = Split(mask, ".")

    For i = 0 To 3
        ' マスクの 0 の部分（ホストビット）が 1 ならネットワークアドレスではない
        ' Not を使うと符号つきの負値になるため 255 - mask でホストビットを求める
        If (CInt(ipParts(i)) And (255 - CInt(maskParts(i)))) <> 0 Then
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
'   範囲外（0未満・32超）の場合は空文字
Function CIDR2Mask(ByVal cidr As Integer) As String
    Dim i As Integer
    Dim mask(3) As Integer
    Dim bits As Integer

    If cidr < 0 Or cidr > 32 Then Exit Function

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
'   形式不正または 1 の連続でない場合は -1
Function Mask2CIDR(ByVal mask As String) As Integer
    Dim parts() As String
    Dim i As Integer
    Dim val As Integer
    Dim cidr As Integer

    ' 形式不正・1の連続でないマスクは -1 を返す
    If Not IsValidSubnetMask(mask) Then
        Mask2CIDR = -1
        Exit Function
    End If

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
