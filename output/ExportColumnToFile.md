### ExportColumnToFile サブルーチン

#### 概要
指定されたワークシートの特定の列のデータをテキストファイルとして書き出すサブルーチンです。列のデータを一度文字列として結合し、指定されたファイルパスに書き込みます。

#### 構文
```vba
Sub ExportColumnToFile(ws As Worksheet, _
                       filePath As String, _
                       Optional colNum As Long = 1, _
                       Optional delimiter As String = vbCrLf)
```

#### パラメータ
| パラメータ名 | 型 | 説明 | 必須/任意 |
| :--- | :--- | :--- | :--- |
| `ws` | `Worksheet` | 処理対象とするワークシートオブジェクト。 | 必須 |
| `filePath` | `String` | データを出力するテキストファイルのフルパス（例: "C:\data\output.txt"）。 | 必須 |
| `colNum` | `Long` | 書き出す列の番号。デフォルトは1（A列）。 | 任意 (Default: 1) |
| `delimiter` | `String` | 列の値と値の間に入れる区切り文字。デフォルトは改行 (`vbCrLf`)。 | 任意 (Default: vbCrLf) |

#### 処理フロー
1.  **値の取得**: 内部で呼び出される`GetColumnValuesAsString`関数を使用して、指定されたワークシートの`colNum`列の全データを、`delimiter`で区切られた単一の文字列`content`として取得します。
2.  **ファイルオープン**: 指定された`filePath`に対して、出力モード（`For Output`）でファイルを開きます。
3.  **書き出し**: 取得した`content`文字列全体をファイルに書き込みます。
4.  **ファイルクローズ**: ファイルを閉じ、リソースを解放します。

#### 備考
*   このサブルーチンは、列のデータを一度文字列として結合してからファイルに書き出すため、大量のデータの場合、メモリ使用量に注意が必要です。
*   `delimiter`をカンマ (`,`) に設定することで、CSV形式での出力が可能です。

