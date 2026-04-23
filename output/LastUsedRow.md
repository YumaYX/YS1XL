### 関数名: LastUsedRow
### 説明
指定されたワークシートの指定された列における、データが入力されている最終行の行番号を返します。
列全体が空の場合、0を返します。

### 構文
```vba
LastUsedRow(ws As Worksheet, Optional col As Long = 1) As Long
```

### パラメータ
| パラメータ名 | 型 | 説明 |
| :--- | :--- | :--- |
| `ws` | `Worksheet` | 最終使用行をチェックする対象のワークシートオブジェクト。 |
| `col` | `Long` | チェックする列の番号。省略された場合、デフォルトで1（A列）が使用されます。 |

### 戻り値
`Long`型：指定された列の最終使用行の行番号を返します。列が完全に空の場合、0を返します。

### 使用例
```vba
' 例: ActiveSheetのB列（列番号2）の最終使用行を取得する
Dim lastRow As Long
lastRow = LastUsedRow(ActiveSheet, 2)
MsgBox "最終使用行: " & lastRow
```

