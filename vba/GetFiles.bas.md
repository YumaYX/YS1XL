# GetFiles

指定フォルダからパターンに一致するファイルのフルパスを返す関数群。
`GetFiles` は対象フォルダのみ、`GetFilesRecursive` はサブフォルダも再帰的に走査する。

含まれる関数: `GetFiles` / `GetFilesRecursive`

---

## GetFiles

- 目的: 指定フォルダ内のパターンに一致するファイルのフルパス一覧を返す。

### 入力

| 引数名 | 型 | 説明 |
|--------|----|------|
| folderPath | String | 対象フォルダパス（末尾に \ がなくても自動補完） |
| pattern | String | ファイル名のワイルドカードパターン（例: *.txt） |

### 出力

- 型: Collection
- 内容: パターンに一致したファイルのフルパスを格納した Collection。
- 対象フォルダが存在しない場合は空の Collection を返す。

### 使用例

```vba
Dim col As Collection
Set col = GetFiles("C:\Temp", "*.txt")

Dim p As Variant
For Each p In col
    Debug.Print p
Next p
```

> 注: `Dir` 命令文を使うため、サブフォルダは対象外。
> 拡張子の指定は大文字小文字を区別しない（Windows の `Dir` の仕様）。

---

## GetFilesRecursive

- 目的: 指定フォルダ以下を再帰的に走査し、パターンに一致するファイルのフルパス一覧を返す。
- 依存: 同モジュール内の Private サブルーチン `GetFilesRecursive_Add` を使用します。

### 入力

| 引数名 | 型 | 説明 |
|--------|----|------|
| folderPath | String | 対象フォルダパス（末尾に \ がなくても自動補完） |
| pattern | String | ファイル名のワイルドカードパターン（例: *.txt） |

### 出力

- 型: Collection
- 内容: サブフォルダを含めてパターンに一致したファイルのフルパスを格納した Collection。
- 対象フォルダが存在しない場合は空の Collection を返す。

### 使用例

```vba
Dim col As Collection
Set col = GetFilesRecursive("C:\Temp", "*.txt")

Dim p As Variant
For Each p In col
    Debug.Print p
Next p
```

> 注: 実装は Private サブルーチン `GetFilesRecursive_Add` を再帰的に呼び出す。
> VBA の `Like` は既定（`Option Compare Binary`）で大文字小文字を区別するため、
> 内部でファイル名・パターンとも `LCase` に変換して比較している。
> そのため `*.txt` は `*.TXT` にも一致する。
