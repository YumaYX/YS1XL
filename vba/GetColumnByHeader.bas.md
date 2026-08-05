# GetColumnByHeader

- 目的: 見出し名から該当する列番号を求める。

# 入力（日本語）

| 引数名        | 型        | 説明                                   |
|---------------|-----------|----------------------------------------|
| ws            | Worksheet | 対象のワークシート                     |
| header        | String    | 探したい見出し名                       |
| rowNum（省略可）| Long    | 見出しが存在する行番号（省略時は 1）   |

# 出力（日本語）

- 型: Long
- 内容: 見出しが見つかった列番号。見つからなければ 0 を返す。

# 使用例

```vba
Dim ws As Worksheet
Set ws = ThisWorkbook.Sheets("Sheet1")

Dim col As Long
col = GetColumnByHeader(ws, "ID")
Debug.Print col
```