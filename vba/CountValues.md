# CountValues

- 目的: 指定したシートの列に含まれる値の出現回数を集計し、`Dictionary` として返します。

# 入力（日本語）

| 引数名 | 型 | 説明 |
|--------|----|------|
| sheetName | String | 必須。集計対象のシート名。 |
| col | Long | オプション。集計対象の列番号（省略時: `1` = A列）。 |

# 出力（日本語）

- 型: `Object` (`Scripting.Dictionary`)
- 内容:
  - `Key`: セルの値
  - `Item`: その値の出現回数

# 使用例

```vb
Dim d As Object
Dim k As Variant

Set d = CountValues("Sheet1")

Debug.Print d("A")   ' A の出現回数

For Each k In d.Keys
    Debug.Print k, d(k)
Next
```