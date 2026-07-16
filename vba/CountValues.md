# CountValues

- 目的: 指定したワークシートの指定列について、各値の出現回数を集計し、`Dictionary` として返します。

# 入力（日本語）

| 引数名 | 型 | 説明 |
|--------|----|------|
| ws | Worksheet | 必須。集計対象のワークシートオブジェクト。 |
| col | Long | オプション。集計対象の列番号。省略時は `1`（A列）。 |

# 出力（日本語）

- 型: `Object`（`Scripting.Dictionary`）
- 内容: 指定列の値ごとの出現回数を格納した `Dictionary` を返します。

| Key | Item |
|-----|------|
| セルの値 | その値の出現回数 |

# 使用例

```vb
Dim d As Object
Dim k As Variant

Set d = CountValues(Worksheets("Sheet1"))

For Each k In d.Keys
    If d(k) > 1 Then
        Debug.Print k
    End If
Next
```

上記の例では、`Sheet1` のA列で2回以上出現する値をイミディエイトウィンドウに出力します。
