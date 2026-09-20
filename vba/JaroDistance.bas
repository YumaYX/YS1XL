Option Explicit

' Jaro距離を返す。0 = 完全一致、1 = 全く異なる
Function JaroDistance(ByVal s1 As String, ByVal s2 As String) As Double

    Dim len1 As Long, len2 As Long
    Dim window As Long
    Dim i As Long, j As Long
    Dim lo As Long, hi As Long
    Dim matches1() As Boolean, matches2() As Boolean
    Dim matchCount As Long
    Dim k As Long
    Dim transposeCnt As Long
    Dim jaro As Double

    len1 = Len(s1)
    len2 = Len(s2)

    ' 両方空、または完全一致なら距離0
    If len1 = 0 And len2 = 0 Then Exit Function
    If s1 = s2 Then Exit Function

    ' マッチ探索窓の半幅
    window = Max(len1, len2) \ 2 - 1
    If window < 0 Then window = 0

    ReDim matches1(1 To len1)
    ReDim matches2(1 To len2)

    ' 探索窓内で一致する文字をマーク
    For i = 1 To len1
        lo = Max(1, i - window)
        hi = Min(i + window, len2)
        For j = lo To hi
            If Not matches2(j) And Mid$(s1, i, 1) = Mid$(s2, j, 1) Then
                matches1(i) = True
                matches2(j) = True
                matchCount = matchCount + 1
                Exit For
            End If
        Next j
    Next i

    ' 共通文字なし → 距離1
    If matchCount = 0 Then
        JaroDistance = 1
        Exit Function
    End If

    ' 転置数（一致文字の並び順の違い）を数える
    k = 1
    For i = 1 To len1
        If matches1(i) Then
            Do While Not matches2(k)
                k = k + 1
            Loop
            If Mid$(s1, i, 1) <> Mid$(s2, k, 1) Then transposeCnt = transposeCnt + 1
            k = k + 1
        End If
    Next i

    ' Jaro類似度 → 距離 = 1 - 類似度
    jaro = (matchCount / len1 + matchCount / len2 + (matchCount - (transposeCnt \ 2)) / matchCount) / 3
    JaroDistance = 1 - jaro

End Function

Private Function Max(a As Long, b As Long) As Long
    If a > b Then Max = a Else Max = b
End Function

Private Function Min(a As Long, b As Long) As Long
    If a < b Then Min = a Else Min = b
End Function

