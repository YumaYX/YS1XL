# CountValues

- 目的: 指定列内の各値の出現回数を集計して Dictionary で返す。

# 入力（日本語）

| 引数名      | 型         | 説明                                   |
|-------------|------------|----------------------------------------|
| ws          | Worksheet  | 対象のワークシート                     |
| col（省略可）| Long       | 集計する列番号（省略時は 1 = A列）     |

# 出力（日本語）

- 型: Object（Scripting.Dictionary）
- 内容: Key=セルの値、Item=出現回数。先頭行から最終使用行までを集計する。

# 使用例

```vba
Dim ws As Worksheet
Set ws = ThisWorkbook.Sheets("Sheet1")

Dim d As Object
Set d = CountValues(ws, 1)

Dim v As Variant
For Each v In d.Keys
    Debug.Print v & " : " & d(v)
Next v
```