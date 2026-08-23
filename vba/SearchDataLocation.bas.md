# StoreDataLocation

- 目的: 2 次元配列から指定列をキーに、データ行番号のインデックスを Dictionary にして返す。

# 入力

| 引数名    | 型      | 説明                                   |
|-----------|---------|----------------------------------------|
| data      | Variant | 2 次元データ配列（例: Range.Value）    |
| keyIndex  | Long    | キーとして使う列番号（1 始まり）       |

# 出力

- 型: Object（Scripting.Dictionary）
- 内容: Key=キー値、Item=該当行番号（重複時は 1 を格納）。

# 使用例

```vba
Dim arr As Variant
arr = ThisWorkbook.Sheets("Sheet1").UsedRange.Value

Dim loc As Object
Set loc = StoreDataLocation(arr, 1)
```

---

# SearchDataLocation

- 目的: CSV ファイルを開き、指定キー列の値から行番号のインデックスを Dictionary にして返す。

# 入力

| 引数名              | 型     | 説明                                   |
|---------------------|--------|----------------------------------------|
| csvFilePath（省略可）| String | 対象 CSV ファイルパス（省略時 "sample.csv"） |
| targetKey（省略可） | String | キーとして使う列の見出し名（省略時 "id"）|

# 出力

- 型: Object（Scripting.Dictionary）
- 内容: Key=キー列の値、Item=該当行番号。処理後はファイルを閉じる。

# 使用例

```vba
Dim loc As Object
Set loc = SearchDataLocation("C:\Temp\sample.csv", "id")
```