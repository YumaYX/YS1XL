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
Function TheHash(ws As Worksheet, keyIndex As Long) As Object
    Dim myHash As Object: Set myHash = CreateObject("Scripting.Dictionary")

    Dim i As Long
    Dim key As String

    For i = 1 To ws.Cells(ws.Rows.Count, keyIndex).End(xlUp).Row
        key = CStr(ws.Cells(i, keyIndex).Value)
        myHash(key) = i
    Next i
    Set TheHash = myHash
End Function
