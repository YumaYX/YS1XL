# GetTimestamp

- 目的: yyyy-mm-dd-HH-MM-ss 形式で現在日時を返す。

# 入力（日本語）

なし

# 出力（日本語）

- 型: String
- 内容: 現在日時を yyyy-mm-dd-HH-MM-ss 形式に整形した文字列。

# 使用例

```vba
Dim ts As String
ts = GetTimestamp()
Debug.Print ts
```