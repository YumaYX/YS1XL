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
'             ※末尾に区切りは付けない
'             ※空列のときは空文字
' 依存     : Application.WorksheetFunction.CountA
'========================================
Function GetColumnValuesAsString(ws As Worksheet, _
                                 Optional colNum As Long = 1, _
                                 Optional delimiter As String = vbCrLf) As String

    Dim lastRow As Long
    Dim result As String
    Dim r As Long
    Dim value As Variant

    GetColumnValuesAsString = ""

    ' 空列のときは End(xlUp) が 1 行目を返すためデータなしとして扱う
    If Application.WorksheetFunction.CountA(ws.Columns(colNum)) = 0 Then Exit Function

    ' 最終行を取得
    lastRow = ws.Cells(ws.Rows.Count, colNum).End(xlUp).Row

    ' 2行目以降は区切りを前に付けることで末尾の区切りを避ける
    For r = 1 To lastRow
        value = ws.Cells(r, colNum).Value
        ' エラー値は文字列連結できないため空文字として扱う
        If IsError(value) Then value = ""
        If r > 1 Then result = result & delimiter
        result = result & value
    Next r

    GetColumnValuesAsString = result
End Function


' Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Sheet1")
' ExportColumnToFile ws, "C:\Temp\out.txt", 1   ' A列を改行区切りで出力
'========================================
' 指定列の値をテキストファイルに書き出す
' ws       : 対象ワークシート
' filePath : 保存先フルパス
' colNum   : 書き出す列番号（省略時1列目）
' delimiter: 値をつなぐ区切り（省略時改行）
' 依存     : GetColumnValuesAsString（本モジュール内）
'========================================
Sub ExportColumnToFile(ws As Worksheet, _
                       filePath As String, _
                       Optional colNum As Long = 1, _
                       Optional delimiter As String = vbCrLf)

    Dim content As String
    Dim fNum As Integer
    Dim isOpen As Boolean

    content = GetColumnValuesAsString(ws, colNum, delimiter)

    ' ファイル書き出し
    fNum = FreeFile
    isOpen = False

    On Error GoTo CloseAndExit

    Open filePath For Output As #fNum
    isOpen = True
    ' 行末の空白改行を付けないため Print の行末を ; で止める
    Print #fNum, content;
    Close #fNum
    isOpen = False
    Exit Sub

CloseAndExit:
    ' エラー時もファイルハンドルを解放する
    On Error Resume Next
    If isOpen Then Close #fNum
    On Error GoTo 0
End Sub
