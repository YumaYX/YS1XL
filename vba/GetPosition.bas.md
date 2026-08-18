# GetPosition

* 頻出:★
* 目的: 指定した文字列を対象ワークシートから完全一致で検索し、最初に一致したセルの座標を取得する。

# 入力（日本語）

| 引数名        | 型         | 説明        |
| ---------- | --------- | --------- |
| ws         | Worksheet | 対象のワークシート |
| searchText | String    | 検索する文字列   |

# 出力（日本語）

* 型: Variant
* 内容: 最初に一致したセルの座標を `Array(Y, X)` 形式で返す。

  * `Y`: 行番号
  * `X`: 列番号
* 見つからない場合は `Array(0, 0)` を返す。
* 文字列は完全一致で判定する。
* 検索は上の行から順番に行う。
* 横方向は、**検索する各行の最終使用列を毎回取得し、その列まで走査する**。
* 同じ文字列が複数存在する場合は、上から走査して最初に一致したセルを返す。

# 使用例

```vba
Dim ws As Worksheet
Set ws = ThisWorkbook.Sheets("Sheet1")

Dim pos As Variant
pos = GetPosition(ws, "商品コード")

Debug.Print "Y: " & pos(0)
Debug.Print "X: " & pos(1)
```

例えば最初に一致したセルが `C5` の場合、

```text
pos(0) = 5
pos(1) = 3
```
