# ExportColumnToFile

- 目的: 指定列の値をテキストファイルに書き出す。

# 入力（日本語）

| 引数名       | 型        | 説明                                   |
|--------------|-----------|----------------------------------------|
| ws           | Worksheet | 対象のワークシート                     |
| filePath     | String    | 保存先のフルパス                       |
| colNum（省略可）| Long    | 書き出す列番号（省略時は 1 = A列）     |
| delimiter（省略可）| String | 値をつなぐ区切り（省略時は改行 vbCrLf） |

# 出力（日本語）

- 型: なし（Sub）
- 内容: GetColumnValuesAsString で列の値を連結し、指定パスへ書き出す。

# 使用例

```vba
Dim ws As Worksheet
Set ws = ThisWorkbook.Sheets("Sheet1")

ExportColumnToFile ws, "C:\Temp\output.txt", 1
```