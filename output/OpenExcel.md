# Function: OpenExcel()

## 概要
ユーザーにファイル選択ダイアログを表示し、選択されたExcelファイル（.xls*）またはCSVファイル（.csv）をワークブックとして開きます。

## 構文
```vba
OpenExcel() As Workbook
```

## 説明
1. `Application.GetOpenFilename` を使用して、ユーザーにファイル選択ダイアログを表示します。
2. ファイルフィルターは「Excelファイル (*.xls*)」と「CSVファイル (*.csv)」に設定されています。
3. ユーザーがキャンセルした場合、`Nothing` を返して処理を終了します。
4. ファイルが正常に選択された場合、`Workbooks.Open` を実行し、開かれたワークブックオブジェクトを返します。

## 戻り値
*   **Workbook:** 選択されたファイルに対応する `Workbook` オブジェクト。
*   **Nothing:** ユーザーがファイル選択をキャンセルした場合。

## 使用例
```vba
Sub TestOpenExcel()
    Dim wb As Workbook
    
    ' ユーザーにファイルを選択させ、開く
    Set wb = OpenExcel()
    
    If Not wb Is Nothing Then
        MsgBox "ワークブックが開かれました: " & wb.Name
    Else
        MsgBox "ファイル選択がキャンセルされました。"
    End If
End Sub
```

