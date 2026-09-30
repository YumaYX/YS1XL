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
    Dim lastCol As Long
    Dim c As Long

    GetColumnByHeader = 0 ' 見つからなければ0

    ' 空行のときは End(xlToLeft) が 1 列目を返すため、そのまま 1 から走査する
    lastCol = ws.Cells(rowNum, ws.Columns.Count).End(xlToLeft).Column

    For c = 1 To lastCol
        If ws.Cells(rowNum, c).Value = header Then
            GetColumnByHeader = c
            Exit Function
        End If
    Next c
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
'             ※空シートのときは Find が Nothing を返すため Array(0, 0)
'----------------------------------------
Function GetPosition(ws As Worksheet, searchText As String) As Variant

    Dim r As Long
    Dim lastRow As Long
    Dim col As Long
    Dim found As Range

    GetPosition = Array(0, 0)

    ' 空シートのときは Find が Nothing を返すため .Row でエラーになる
    Set found = ws.Cells.Find("*", _
                              SearchOrder:=xlByRows, _
                              SearchDirection:=xlPrevious)
    If found Is Nothing Then Exit Function

    lastRow = found.Row

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
'                IDは文字列化して比較するため、数値/文字列の型差は吸収される
' 依存         : GetColumnByHeader（本モジュール内）
'                Application.WorksheetFunction.CountA
'========================================
Function GetValueByID(ws As Worksheet, _
                             idHeader As String, _
                             idValue As Variant, _
                             targetHeader As String, _
                             Optional headerRow As Long = 1) As Variant

    Dim idCol     As Long
    Dim targetCol As Long
    Dim lastRow   As Long
    Dim r         As Long
    Dim cellValue As Variant
    Dim found     As Variant

    GetValueByID = ""

    idCol     = GetColumnByHeader(ws, idHeader, headerRow)
    targetCol = GetColumnByHeader(ws, targetHeader, headerRow)
    If idCol = 0 Or targetCol = 0 Then Exit Function

    ' 空列のときは End(xlUp) が 1 行目を返すためデータなしとして扱う
    If Application.WorksheetFunction.CountA(ws.Columns(idCol)) = 0 Then Exit Function

    lastRow = ws.Cells(ws.Rows.Count, idCol).End(xlUp).Row

    For r = headerRow + 1 To lastRow
        cellValue = ws.Cells(r, idCol).Value
        If Not IsError(cellValue) Then ' エラー値回避
            ' Variant 同士でも型の差で比較が False になるため文字列化して比較する
            If CStr(cellValue) = CStr(idValue) Then
                found = ws.Cells(r, targetCol).Value
                If Not IsError(found) Then GetValueByID = found
                Exit Function
            End If
        End If
    Next r
End Function
