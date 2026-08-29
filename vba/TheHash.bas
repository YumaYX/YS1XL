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
