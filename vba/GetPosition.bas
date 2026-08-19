'######### GetPosition
' Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Sheet1")
' Dim pos As Variant: pos = GetPosition(ws, "商品コード")
' Debug.Print pos(0) & "," & pos(1)   ' 例: "5,3"（C5）
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

