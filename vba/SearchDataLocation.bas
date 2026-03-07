'######### SearchDataLocation
Function StoreDataLocation(data As Variant, keyIndex As Long) As Object

    Dim dataLocation As Object
    Set dataLocation = CreateObject("Scripting.Dictionary")

    Dim i As Long
    Dim key As String
    Dim idx As Long

    For i = 1 To UBound(data, 1)
        key = CStr(data(i, keyIndex))
        dataLocation(key) = IIf(dataLocation.Exists(key), 1, i)
    Next i
    Set StoreDataLocation = dataLocation
End Function


Function SearchDataLocation(Optional csvFilePath As String = "sample.csv", _
                            Optional targetKey As String = "id") As Object

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim dataArr As Variant

    Set wb = Workbooks.Open(csvFilePath)
    Set ws = wb.Sheets(1)

    dataArr = ws.UsedRange.Value

    Dim headersHash As Object
    Set headersHash = CreateObject("Scripting.Dictionary")

    Dim col As Long
    Dim lastCol As Long
    lastCol = UBound(dataArr, 2)

    For col = 1 To lastCol
        headersHash(CStr(dataArr(1, col))) = col
    Next col

    Dim keyColumnIndexNumber As Long
    keyColumnIndexNumber = headersHash(targetKey)

    Set SearchDataLocation = StoreDataLocation(dataArr, keyColumnIndexNumber)

    wb.Close False

End Function