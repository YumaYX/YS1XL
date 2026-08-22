'######### ExportColumn
' 列の値を取得・ファイル出力する統合モジュール
' 含まれる関数: GetColumnValuesAsString / ExportColumnToFile
'========================================
' GetColumnValuesAsString - 指定列の値を区切り文字で連結した文字列を返す
' ExportColumnToFile      - 指定列の値をテキストファイルに書き出す
'========================================

' Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Sheet1")
' Dim s As String: s = GetColumnValuesAsString(ws, 1, " / ")
' Debug.Print s    ' "A1 / A2 / A3 ..." （改行区切りのときは vbCrLf）
'========================================
' 指定列の値を文字列で返す
' ws        : 対象ワークシート
' colNum    : 取得する列番号（省略時1列目）
' delimiter : 値をつなぐ区切り（省略時 vbCrLf で改行）
' 戻り値    : 列の値をつなげた文字列
'========================================
Function GetColumnValuesAsString(ws As Worksheet, _
                                 Optional colNum As Long = 1, _
                                 Optional delimiter As String = vbCrLf) As String
    ' 最終行を取得
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, colNum).End(xlUp).Row
    Dim result As String: result = ""

    Dim r As Long: For r = 1 To lastRow
        result = result & ws.Cells(r, colNum).Value & delimiter
    Next r
    GetColumnValuesAsString = result
End Function


' Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Sheet1")
' ExportColumnToFile ws, "C:\Temp\out.txt", 1   ' A列を改行区切りで出力
'========================================
' 指定列の値をテキストファイルに書き出す
' ws       : 対象ワークシート
' colNum   : 書き出す列番号（省略時1列目）
' filePath : 保存先フルパス
' delimiter: 値をつなぐ区切り（省略時改行）
' 依存     : GetColumnValuesAsString（本モジュール内）
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
