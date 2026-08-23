# LastUsedRow

- 目的: 指定列の最終使用行番号を返す（値がなければ 0）。

# 入力

| 引数名        | 型         | 説明                                   |
|---------------|------------|----------------------------------------|
| ws            | Worksheet  | 対象のワークシート                     |
| col（省略可） | Long       | 対象列番号（省略時は 1 = A列）         |

# 出力

- 型: Long
- 内容: 指定列の最終使用行番号。列が空なら 0。

# 使用例

```vba
Dim ws As Worksheet
Set ws = ThisWorkbook.Sheets("Sheet1")

Dim lastRow As Long
lastRow = LastUsedRow(ws, 1)
Debug.Print lastRow
```