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
'     Key  : セルの値を文字列化したもの
'     Item : 出現回数
'   ※空列のときは空の Dictionary を返す
'   ※エラー値（#N/A など）は集計対象から除外する
' 依存:
'   Application.WorksheetFunction.CountA
'========================================

Function CountValues(ws As Worksheet, Optional col As Long = 1) As Object

    Dim lastRow As Long
    Dim c As Range
    Dim d As Object
    Dim key As String

    Set d = CreateObject("Scripting.Dictionary")

    ' 空列のときは End(xlUp) が 1 行目を返すためデータなしとして扱う
    If Application.WorksheetFunction.CountA(ws.Columns(col)) = 0 Then
        Set CountValues = d
        Exit Function
    End If

    lastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row

    For Each c In ws.Range(ws.Cells(1, col), ws.Cells(lastRow, col))
        ' エラー値は Dictionary のキーにできないため除外する
        If Not IsError(c.Value) Then
            key = CStr(c.Value)
            If d.Exists(key) Then
                d(key) = d(key) + 1
            Else
                d.Add key, 1
            End If
        End If
    Next

    Set CountValues = d

End Function
