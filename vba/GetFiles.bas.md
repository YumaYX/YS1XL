# GetFiles

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
Set col = GetFiles("C:\Temp", "*.txt")

Dim p As Variant
For Each p In col
    Debug.Print p
Next p
```

---

# GetFilesRecursive

- 目的: 指定フォルダ以下を再帰的に走査し、パターンに一致するファイルのフルパス一覧を返す。

## 入力

| 引数名 | 型 | 説明 |
|--------|----|------|
| folderPath | String | 対象フォルダパス（末尾に \ がなくても自動補完） |
| pattern | String | ファイル名のワイルドカードパターン（例: *.txt） |

## 出力

- 型: Collection
- 内容: サブフォルダを含めてパターンに一致したファイルのフルパスを格納した Collection。

## 使用例

```vba
Dim col As Collection
Set col = GetFilesRecursive("C:\Temp", "*.txt")

Dim p As Variant
For Each p In col
    Debug.Print p
Next p
```

> 注: 実装は後述の Private サブルーチン `GetFilesRecursive_Add` を再帰的に呼び出す。対象フォルダが存在しない場合は空の Collection を返す。

---

# GetFilesRecursive_Add (Private)

- 目的: `GetFilesRecursive` から呼び出される再帰処理。指定フォルダ内のファイルとサブフォルダを走査し、パターンに一致するファイルのフルパスを Collection に追加する。

## 入力

| 引数名 | 型 | 説明 |
|--------|----|------|
| folder | Object | 走査対象の FSO Folder オブジェクト |
| pattern | String | ファイル名のワイルドカードパターン（例: *.txt） |
| col | Collection | 一致したファイルのフルパスを追加する Collection（ByRef） |

## 出力

なし（`col` に直接追加される）。
