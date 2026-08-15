# GetValueByID

- 頻出:★
- 目的: ID をキーに、対象ワークシートから該当するセルの値を取得する。

# 入力（日本語）

| 引数名            | 型        | 説明                                   |
|-------------------|-----------|----------------------------------------|
| ws                | Worksheet | 対象のワークシート                     |
| idHeader          | String    | ID 列の見出し名                        |
| idValue           | Variant   | 検索する ID の値                      |
| targetHeader      | String    | 取得したい列の見出し名                 |
| headerRow（省略可）| Long    | 見出し行番号（省略時は 1）             |

# 出力（日本語）

- 型: Variant
- 内容: 該当セルの値（ID または対象列が見つからなければ "" を返す）。エラー値を回避。

# 使用例

```vba
Dim ws As Worksheet
Set ws = ThisWorkbook.Sheets("Sheet1")

Dim v As Variant
v = GetValueByID(ws, "ID", "A001", "名前")
Debug.Print v
```