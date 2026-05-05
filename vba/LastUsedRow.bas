'######### LastUsedRow
Function LastUsedRow(ws As Worksheet, Optional col As Long = 1) As Long
    With ws
        If Application.WorksheetFunction.CountA(.Columns(col)) = 0 Then
            LastUsedRow = 0
        Else
            LastUsedRow = .Cells(.Rows.Count, col).End(xlUp).Row
        End If
    End With
End Function

