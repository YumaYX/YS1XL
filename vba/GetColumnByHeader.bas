'######### GetColumnByHeader
'========================================
' 見出し名から列番号を探す
' ws      : 対象ワークシート
' header  : 探したい見出し名
' rowNum  : 見出しがある行番号（通常1行目）
' 戻り値  : 列番号（見つからなければ0）
'========================================
Function GetColumnByHeader(ws As Worksheet, header As String, Optional rowNum As Long = 1) As Long
    Dim lastCol As Long: lastCol = ws.Cells(rowNum, ws.Columns.Count).End(xlToLeft).Column
    Dim c As Long: For c = 1 To lastCol
        GetColumnByHeader = c
        If ws.Cells(rowNum, c).Value = header Then Exit Function
    Next c
    GetColumnByHeader = 0 ' 見つからなければ0
End Function

