# SelectFolder

目的: ユーザーが選択したフォルダのパスを取得する。


# 入力

なし。

# 出力

型: String

内容: ユーザーが選択したフォルダのパス。キャンセルした場合は空文字列を返す。


# 使用例

```
Dim folderPath As String

folderPath = SelectFolder()

If folderPath <> "" Then
    MsgBox "選択したフォルダ: " & folderPath
End If
```