'######### OpenExcel
'# Dim wb As Workbook: Set wb = OpenExcel()

Function OpenExcel() As Workbook
    Dim filename As Variant
    filename = Application.GetOpenFilename( _
        FileFilter:="Excelファイル (*.xls*),*.xls*,CSVファイル (*.csv),*.csv")

    If filename = False Then
        Set OpenExcel = Nothing
        Exit Function
    End If

    Set OpenExcel = Workbooks.Open(filename)
End Function

