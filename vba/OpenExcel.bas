' Dim wb As Workbook: Set wb = OpenExcel()
' If wb Is Nothing Then Exit Sub   ' キャンセル時は Nothing が返る

Function OpenExcel() As Workbook
    Dim filename As Variant

    OpenExcel = Nothing

    filename = Application.GetOpenFilename( _
        FileFilter:="Excelファイル (*.xls*),*.xls*,CSVファイル (*.csv),*.csv")

    ' キャンセル時は Boolean の False が返るため VarType で判定する
    If VarType(filename) = vbBoolean Then Exit Function
    If Len(filename) = 0 Then Exit Function

    Set OpenExcel = Workbooks.Open(filename)
End Function
