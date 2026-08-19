'######### LastUsedRow
' Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Sheet1")
' Dim n As Long: n = LastUsedRow(ws)      ' A列の最終使用行
' Debug.Print n                           ' 例: 100
Function LastUsedRow(ws As Worksheet, Optional col As Long = 1) As Long
    With ws
        If Application.WorksheetFunction.CountA(.Columns(col)) = 0 Then
            LastUsedRow = 0
        Else
            LastUsedRow = .Cells(.Rows.Count, col).End(xlUp).Row
        End If
    End With
End Function


