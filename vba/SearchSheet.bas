'######### SearchSheet
' ワークシート内の検索・値取得 統合モジュール
' 含まれる関数: GetColumnByHeader / GetPosition / GetValueByID
'========================================
' GetColumnByHeader - 見出し名から列番号を探す
' GetPosition       - 文字列を検索し最初に一致したセル座標を返す
' GetValueByID      - ID列の値から対象列の値を取得する
'========================================

' Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Sheet1")
' Dim col As Long: col = GetColumnByHeader(ws, "ID")
' Debug.Print col              ' 見出し"ID"の列番号（例: 2）
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


' Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Sheet1")
' Dim pos As Variant: pos = GetPosition(ws, "商品コード")
' Debug.Print pos(0) & "," & pos(1)   ' 例: "5,3"（C5）
'----------------------------------------
' 指定した文字列を検索し、最初に一致したセル座標を返す
' ws        : 対象ワークシート
' searchText: 検索する文字列（完全一致）
' 戻り値    : Array(Y, X) 見つからなければ Array(0, 0)
'             Y: 行番号 / X: 列番号
' 依存      : GetColumnByHeader（本モジュール内）
Function GetPosition(ws As Worksheet, searchText As String) As Variant

    Dim r As Long
    Dim lastRow As Long
    Dim col As Long

    GetPosition = Array(0, 0)

    lastRow = ws.Cells.Find("*", _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlPrevious).Row

    For r = 1 To lastRow

        'この行の最終列まで検索
        col = GetColumnByHeader(ws, searchText, r)

        If col > 0 Then
            GetPosition = Array(r, col)
            Exit Function
        End If

    Next r

End Function


' Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Sheet1")
' Dim v As Variant: v = GetValueByID(ws, "ID", 123, "名前")
' Debug.Print v    ' ID=123 の行の"名前"列の値（見つからなければ ""）
'========================================
' IDから値取得（ID列・取得列は自動検索）
' ws           : 対象ワークシート
' idHeader     : ID列の見出し名
' idValue      : 検索するID
' targetHeader : 取得したい列の見出し名
' headerRow    : 見出し行番号（省略可、通常1）
' 戻り値       : 該当セルの値（見つからなければ""）
' 依存         : GetColumnByHeader（本モジュール内）
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
