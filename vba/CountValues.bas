' Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Sheet1")
' Dim d As Object: Set d = CountValues(ws, 1)
' Debug.Print d("りんご")   ' "りんご" の出現回数
'========================================
' 列内の値の出現回数を集計
'----------------------------------------
' 引数:
'   ws  - Worksheetオブジェクト
'   col - 列番号（省略時: A列）
'
' 戻り値:
'   Dictionary
'     Key  : セルの値
'     Item : 出現回数
'========================================

Function CountValues(ws As Worksheet, Optional col As Long = 1) As Object

    Dim lastRow As Long
    Dim c As Range
    Dim d As Object

    lastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row

    Set d = CreateObject("Scripting.Dictionary")

    For Each c In ws.Range(ws.Cells(1, col), ws.Cells(lastRow, col))
        d(c.Value) = d(c.Value) + 1
    Next

    Set CountValues = d

End Function




