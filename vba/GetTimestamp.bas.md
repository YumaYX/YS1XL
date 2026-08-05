# GetTimestamp

- 目的: yyyy-mm-dd-HH-MM-ss 形式で現在時刻を返す。

# 入力（日本語）

| 引数名 | 型 | 説明 |
|--------|----|------|

# 出力（日本語）

- 型: String
- 内容: 現在日時を yyyy-mm-dd-HH-MM-ss 形式で返す。

# 使用例

```vba
Dim t As String
t = GetTimestamp()
Debug.Print t
```