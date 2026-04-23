### 関数名: GetColumnValuesAsString
### 説明:
指定されたワークシートの指定列の全データを文字列として結合して返します。

### 構文:
```vba
GetColumnValuesAsString(ws As Worksheet, Optional colNum As Long = 1, Optional delimiter As String = vbCrLf) As String
```

### 引数 (Parameters):
| 引数名 | 型 | 説明 | 必須/任意 | デフォルト値 |
| :--- | :--- | :--- | :--- | :--- |
| `ws` | `Worksheet` | 値を取得する対象のワークシート。 | 必須 | - |
| `colNum` | `Long` | 値を取得する列番号。 | 任意 | 1 (A列) |
| `delimiter` | `String` | 各セルの値を結合する際の区切り文字。 | 任意 | `vbCrLf` (改行) |

### 戻り値 (Return Value):
指定された列の全セル（1行目から最終行まで）の値が、`delimiter`で結合された単一の文字列。

### 動作概要:
1.  指定されたワークシート(`ws`)の、指定された列(`colNum`)の最終行を特定します。
2.  1行目から最終行までをループし、各セルの値を取得します。
3.  取得した値を順次、`delimiter`で結合し、最終的な文字列として返します。

