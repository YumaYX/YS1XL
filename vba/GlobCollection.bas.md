# GlobCollection

- 目的: 指定フォルダ内のパターンに一致するファイルのフルパス一覧を返す。

## 入力

| 引数名 | 型 | 説明 |
|--------|----|------|
| folderPath | String | 対象フォルダパス（末尾に \ がなくても自動補完） |
| pattern | String | ファイル名のワイルドカードパターン（例: *.txt） |

## 出力

- 型: Collection
- 内容: パターンに一致したファイルのフルパスを格納した Collection。

## 使用例

```vba
Dim col As Collection
Set col = GlobCollection("C:\Temp", "*.txt")

Dim p As Variant
For Each p In col
    Debug.Print p
Next p
```
