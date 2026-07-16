'######### CountValues
'========================================
' 列内の値の出現回数を集計
'----------------------------------------
' 引数:
'   sheetName - シート名
'   col       - 列番号（省略時: A列）
'
' 戻り値:
'   Dictionary
'     Key  : セルの値
'     Item : 出現回数
'========================================
Function CountValues(sheetName As String, Optional col As Long = 1) As Object

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim c As Range
    Dim d As Object

    Set ws = Worksheets(sheetName)
    lastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row

    Set d = CreateObject("Scripting.Dictionary")

    For Each c In ws.Range(ws.Cells(1, col), ws.Cells(lastRow, col))
        d(c.Value) = d(c.Value) + 1
    Next

    Set CountValues = d

End Function

' for use
'For Each k In d.Keys
'    If d(k) > 1 Then
'        Debug.Print k
'    End If
'Next

