# ReadUtf8Text

- 目的: UTF-8 エンコードのテキストファイルを読み込み、その内容を返す。

# 入力（日本語）

| 引数名   | 型     | 説明             |
|----------|--------|------------------|
| filePath | String | 読み込むファイルのフルパス |

# 出力（日本語）

- 型: String
- 内容: ファイルのテキスト内容。

# 使用例

```vba
Dim t As String
t = ReadUtf8Text("C:\Temp\data.txt")
Debug.Print t
```