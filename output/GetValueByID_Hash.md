# 関数名: GetValueByID_Hash

## 概要
指定されたワークシートにおいて、特定のヘッダー名（ID列）を基準として検索値（ID）を探し、一致した行の別のヘッダー名（取得列）に対応する値を取得します。列番号を固定する必要がなく、ヘッダー名のみで柔軟なデータ検索（ハッシュ検索）が可能です。

## 構文
```vba
Function GetValueByID_Hash(ws As Worksheet, idHeader As String, idValue As Variant, targetHeader As String, Optional headerRow As Long = 1) As Variant
```

## 引数 (Parameters)
| 引数名 | 型 | 説明 | 必須/任意 |
| :--- | :--- | :--- | :--- |
| `ws` | `Worksheet` | 検索対象とするワークシートオブジェクト。 | 必須 |
| `idHeader` | `String` | IDとして使用する列の見出し名（ヘッダー名）。 | 必須 |
| `idValue` | `Variant` | 検索したいIDの値。 | 必須 |
| `targetHeader` | `String` | 取得したい値が含まれる列の見出し名（ヘッダー名）。 | 必須 |
| `headerRow` | `Long` | 見出し行の行番号。指定しない場合、デフォルトで1行目を使用します。 | 任意 (デフォルト: 1) |

## 戻り値 (Return Value)
*   **型:** `Variant`
*   **内容:** `idValue`に一致した行の`targetHeader`に対応するセルの値が返されます。
*   **失敗時:** 該当するIDが見つからない場合、または指定されたヘッダー列が存在しない場合は、空文字列 (`""`) が返されます。

## 処理フロー
1.  `idHeader`と`targetHeader`を用いて、それぞれ対応する列番号をワークシートから自動検索します。
2.  検索対象の最終行を特定します。
3.  ヘッダー行の次の行から最終行までを順にループ処理します。
4.  ループ中に、ID列の値が`idValue`と一致するかをチェックします。
5.  一致した場合、対応する行の取得列の値を返して処理を終了します。

