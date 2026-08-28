# TheHash

- 目的: ワークシートの指定列をキーとした連想配列（ディクショナリ）を作成する

# 入力

| 引数名 | 型 | 説明 |
|--------|----|------|
| ws | Worksheet | 対象のワークシート |
| keyIndex | Long | キーとする列番号 |

# 出力

- 型: Object（Scripting.Dictionary）
- 内容: セルの値をキー、行番号を値とする連想配列

# 使用例

```vba
Dim myHash As Object
Set myHash = TheHash(Sheets("Sheet1"), 1)
Debug.Print myHash("someKey")  ' キーに対応する行番号が返る
```
