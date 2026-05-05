'######### ExportColumnToFile
'========================================
' 指定列の値をテキストファイルに書き出す
' ws       : 対象ワークシート
' colNum   : 書き出す列番号（省略時1列目）
' filePath : 保存先フルパス
' delimiter: 値をつなぐ区切り（省略時改行）
'========================================
Sub ExportColumnToFile(ws As Worksheet, _
                       filePath As String, _
                       Optional colNum As Long = 1, _
                       Optional delimiter As String = vbCrLf)
    Dim content As String: content = GetColumnValuesAsString(ws, colNum, delimiter)
    ' ファイル書き出し
    Dim fNum As Integer: fNum = FreeFile
    Open filePath For Output As #fNum
    Print #fNum, content
    Close #fNum
End Sub

