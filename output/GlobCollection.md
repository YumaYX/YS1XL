# 関数名: GlobCollection

## 概要
指定されたフォルダパスとファイルパターンに基づいて、そのフォルダ内にあるすべてのファイルパスを検索し、それらを`Collection`オブジェクトとして返します。これは、指定されたディレクトリ内の複数のファイルを効率的にリストアップするために使用されます。

## 構文
```vba
Function GlobCollection(folderPath As String, pattern As String) As Collection
```

## パラメータ (Parameters)
| パラメータ名 | 型 | 説明 |
| :--- | :--- | :--- |
| `folderPath` | `String` | ファイル検索を開始するフォルダの完全なパス。パスの末尾にバックスラッシュ（`\`）がない場合、自動的に追加されます。 |
| `pattern` | `String` | 検索するファイル名パターン。ワイルドカード（`*`や`?`）を使用できます（例: `*.txt`）。 |

## 戻り値 (Return Value)
`Collection`
検索されたすべてのファイルのフルパス（例: `C:\path\to\file.txt`）を格納したコレクションオブジェクトを返します。ファイルが見つからない場合、空のコレクションが返されます。

## 処理の流れ
1.  引数として渡された`folderPath`が末尾に`\`を持たない場合、自動的に追加してパスを修正します。
2.  `Dir`関数を使用して、指定された`folderPath`と`pattern`に一致する最初のファイル名を取得します。
3.  `Do While`ループを使用し、ファイル名が空文字列（`""`）になるまで以下の処理を繰り返します。
    *   現在のファイル名とフォルダパスを結合し、完全なファイルパスを`Collection`に追加します。
    *   `Dir()`関数を引数なしで呼び出し、次のファイル名を取得します。
4.  ループが終了した後、構築された`Collection`を関数の戻り値として返します。

## 使用例 (Conceptual Example)
```vba
' 例: "C:\Data"フォルダ内のすべてのPDFファイルを取得する
Dim folder As String
Dim pattern As String
Dim fileCollection As Collection

folder = "C:\Data"
pattern = "*.pdf"

Set fileCollection = GlobCollection(folder, pattern)

' コレクション内の要素をループ処理する
If fileCollection.Count > 0 Then
    For Each item In fileCollection
        Debug.Print item ' 各ファイルのフルパスが出力される
    Next item
End If
```

