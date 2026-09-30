' 使用例:
'   Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Sheet1")
'   Dim h As Object: Set h = TheHash(ws, 1)
'   Debug.Print h("りんご")          ' 値"りんご"が最後に現れた行番号
' 引数:
'   ws       - 対象ワークシート
'   keyIndex - キーとして使う列番号
' 戻り値:
'   Scripting.Dictionary
'     Key  : セルの値（文字列化）
'     Item : その値が最後に現れた行番号
'   ※同じ値が複数ある場合は最後の行番号で上書きされる
'   ※空列のときは空の Dictionary を返す
'   ※エラー値（#N/A など）は除外される
' 依存:
'   Application.WorksheetFunction.CountA
Function TheHash(ws As Worksheet, keyIndex As Long) As Object
    Dim myHash As Object: Set myHash = CreateObject("Scripting.Dictionary")

    Dim i As Long
    Dim key As String
    Dim value As Variant
    Dim lastRow As Long

    ' 空列のときは End(xlUp) が 1 行目を返すためデータなしとして扱う
    If Application.WorksheetFunction.CountA(ws.Columns(keyIndex)) = 0 Then
        Set TheHash = myHash
        Exit Function
    End If

    lastRow = ws.Cells(ws.Rows.Count, keyIndex).End(xlUp).Row

    For i = 1 To lastRow
        value = ws.Cells(i, keyIndex).Value
        ' エラー値は CStr できないため除外する
        If Not IsError(value) Then
            key = CStr(value)
            myHash(key) = i
        End If
    Next i

    Set TheHash = myHash
End Function
