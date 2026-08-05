# GetColumnValuesAsString

- 目的: 指定列の値を区切り文字で連結した文字列を返す。

# 入力（日本語）

| 引数名          | 型        | 説明                                   |
|-----------------|-----------|----------------------------------------|
| ws              | Worksheet | 対象のワークシート                     |
| colNum（省略可）| Long    | 取得する列番号（省略時は 1 = A列）     |
| delimiter（省略可）| String | 値をつなぐ区切り（省略時は改行 vbCrLf）|

# 出力（日本語）

- 型: String
- 内容: 先頭行から最終使用行までの各値を区切りでつないだ文字列。

# 使用例

```vba
Dim ws As Worksheet
Set ws = ThisWorkbook.Sheets("Sheet1")

Dim s As String
s = GetColumnValuesAsString(ws, 1, ",")
Debug.Print s
```