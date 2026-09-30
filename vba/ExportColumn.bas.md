# ExportColumn

列の値を取得・ファイル出力する統合モジュール。

含まれる関数: `GetColumnValuesAsString` / `ExportColumnToFile`

---

## GetColumnValuesAsString

- 目的: 指定列の値を区切り文字で連結した文字列を返す。

### 入力

| 引数名 | 型 | 説明 |
|--------|----|------|
| ws | Worksheet | 対象のワークシート |
| colNum（省略可） | Long | 取得する列番号（省略時は 1 = A列） |
| delimiter（省略可） | String | 値をつなぐ区切り（省略時は改行 vbCrLf） |

### 出力

- 型: String
- 内容: 先頭行から最終使用行までの各値を区切りでつないだ文字列。
- 末尾に区切りは付かない（A列が `A1` / `A2` / `A3` なら `"A1<区切り>A2<区切り>A3"`）。
- 空列のときは空文字 `""` を返す。
- エラー値（`#N/A` など）は空文字として扱って連結する。

### 使用例

```vba
Dim ws As Worksheet
Set ws = ThisWorkbook.Sheets("Sheet1")

Dim s As String
s = GetColumnValuesAsString(ws, 1, " / ")
Debug.Print s
```

---

## ExportColumnToFile

- 目的: 指定列の値をテキストファイルに書き出す。
- 依存: 同モジュール内の `GetColumnValuesAsString` を使用する。

### 入力

| 引数名 | 型 | 説明 |
|--------|----|------|
| ws | Worksheet | 対象のワークシート |
| filePath | String | 保存先のフルパス |
| colNum（省略可） | Long | 書き出す列番号（省略時は 1 = A列） |
| delimiter（省略可） | String | 値をつなぐ区切り（省略時は改行 vbCrLf） |

### 出力

- 型: なし（Sub）
- 内容: `GetColumnValuesAsString` で列の値を連結し、指定パスへ書き出す。
- 行末の空白改行は付けない（`Print #fNum, content;` の末尾の `;` による）。
- 書き込み中にエラーが発生した場合はファイルハンドルを解放した上で処理を終える。

### 使用例

```vba
Dim ws As Worksheet
Set ws = ThisWorkbook.Sheets("Sheet1")

ExportColumnToFile ws, "C:\Temp\output.txt", 1
```
