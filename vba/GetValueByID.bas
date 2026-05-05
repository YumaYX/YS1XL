'######### GetValueByID
'========================================
' IDから値取得（ID列・取得列は自動検索）
' ws           : 対象ワークシート
' idHeader     : ID列の見出し名
' idValue      : 検索するID
' targetHeader : 取得したい列の見出し名
' headerRow    : 見出し行番号（省略可、通常1）
' 戻り値       : 該当セルの値（見つからなければ""）
'========================================
Function GetValueByID(ws As Worksheet, _
                             idHeader As String, _
                             idValue As Variant, _
                             targetHeader As String, _
                             Optional headerRow As Long = 1) As Variant
    GetValueByID = ""
    
    Dim idCol     As Long: idCol     = GetColumnByHeader(ws, idHeader, headerRow)
    Dim targetCol As Long: targetCol = GetColumnByHeader(ws, targetHeader, headerRow)
    If idCol = 0 Or targetCol = 0 Then Exit Function    

    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, idCol).End(xlUp).Row

    Dim r As Long: For r = headerRow + 1 To lastRow
        If Not IsError(ws.Cells(r, idCol).Value) Then ' エラー値回避
            If ws.Cells(r, idCol).Value = idValue Then
                GetValueByID = ws.Cells(r, targetCol).Value
                Exit Function
            End If
        End If
    Next r
End Function

