### 関数名: GetValueByID
### 概要:
指定されたワークシートにおいて、ID列（`idHeader`）を検索し、指定されたID値（`idValue`）と一致する行を見つけます。一致した行の、指定されたターゲット列（`targetHeader`）の値を取得して返します。

### 構文:
```vba
GetValueByID(ws As Worksheet, idHeader As String, idValue As Variant, targetHeader As String, Optional headerRow As Long = 1) As Variant
```

### 引数 (Parameters):
| 引数名 | 型 | 説明 | 必須/任意 |
| :--- | :--- | :--- | :--- |
| `ws` | `Worksheet` | 処理対象となるワークシートオブジェクト。 | 必須 |
| `idHeader` | `String` | ID値が格納されている列の見出し名（ヘッダー名）。 | 必須 |
| `idValue` | `Variant` | 検索したいIDの値。 | 必須 |
| `targetHeader` | `String` | 取得したい値が格納されている列の見出し名（ヘッダー名）。 | 必須 |
| `headerRow` | `Long` | 見出し行の行番号。デフォルトは1。 | 任意 |

### 戻り値 (Return Value):
*   **`Variant`**: IDが一致した行の、ターゲット列の値が返されます。
*   **`""`**: IDが見つからなかった場合、または指定された列が見つからない場合に空文字列が返されます。

### 処理フロー:
1.  `GetColumnByHeader`関数を使用して、`idHeader`と`targetHeader`に対応する列番号を取得します。
2.  ID列の最終行を特定します。
3.  ヘッダー行の次の行から最終行までループ処理を行います。
4.  現在の行のID列の値が`idValue`と一致するかどうかをチェックします。
5.  一致した場合、対応する行のターゲット列の値を返して処理を終了します。
6.  ループが最後まで実行され、一致するIDが見つからなかった場合は、空の値を返します。

