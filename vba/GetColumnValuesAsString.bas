'######### GetColumnValuesAsString
'========================================
' 指定列の値を文字列で返す
' ws        : 対象ワークシート
' colNum    : 取得する列番号（省略時1列目）
' delimiter : 値をつなぐ区切り（省略時 vbCrLf で改行）
' 戻り値    : 列の値をつなげた文字列
'========================================
Function GetColumnValuesAsString(ws As Worksheet, _
                                 Optional colNum As Long = 1, _
                                 Optional delimiter As String = vbCrLf) As String
    ' 最終行を取得
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, colNum).End(xlUp).Row
    Dim result As String: result = ""

    Dim r As Long: For r = 1 To lastRow
        result = result & ws.Cells(r, colNum).Value & delimiter
    Next r
    GetColumnValuesAsString = result
End Function

