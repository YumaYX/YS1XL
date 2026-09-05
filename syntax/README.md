# Excel VBA チートシート

## 数秒のウェイト

```vb
Application.Wait Now + TimeValue("00:00:02")
```

## セル操作

- Cells
  `Cells(Y↓,X→)`
- Range/Cell/Table
  `Range("cell / cell name / table")`
- エクセルのセルは**1スタート**。

### セル内容削除

```vb
worksheet.Cells.clear
```

### セル自体の削除（詰める）

```vb
Selection.Delete
```

## 変数宣言

- 型を宣言しない変数に値を投げるとエラーになる。**事前に型を宣言する。**
- active の変数は使わない方がいい。

```vb
' 型宣言
Dim str As String
Dim l As Long

' 1行で宣言+代入
Dim str As String: str = "abcd"
Dim l As Long: l = 1
```

## 制御構文

### If

```vb
If 条件 Then
    '処理
ElseIf 条件 Then
    '処理
Else
    '処理
End if
```

### For

```vb
Dim i As Integer
For i = 1 To 10
    Debug.Print(i)
Next i
```

### For Each

```vb
Dim f As Variant
For Each f In arr
    Debug.Print f
Next f
```

## 関数定義

### 値渡し `ByVal`

```vb
Sub f1 (ByVal str As String)
  'Debug.Print(str)
End Sub

Call f1("hello")
```

### 参照渡し `ByRef`（省略時のデフォルト）

```vb
Sub f2 (ByRef str As String)
  'Debug.Print(str)
End Sub

str="hello"
Call f2(str)
```

### 返り値ありパターン

`as 型`をつけ、最後に**関数名へ代入**。`Set`がいるやつは`Set`。

```vb
Sub f1 () as String
  f1 = "hello"
End Sub
```

### 注意・その他

- `Set`で宣言するものは、`Set`をつけて投げ、`Set`で受け取る。
- 引数の型:`Object`もある。
- Workbookのとき、引数の型`ByRef`, `ByVal`は書かなくてもいい。
  - `Sub SampleSub(wb As Workbook)`
  - デフォルトが`ByRef`
- 明らかに型が合致している場合でも、型が合わないエラーが出る。

## Array

### 固定配列

```vb
Dim arr As Variant
arr = Array("a", "b", "c")

Dim f As Variant
For Each f In arr
    Debug.Print f
Next f
```

### 可変長（動的）配列 `ReDim`

`ReDim`で大きさを再定義できる。`Preserve`をつけると既存の値を保持しつつ拡張できる。

```vb
Dim arr() As String
ReDim arr(0)
arr(0) = "a"

' 末尾に追加（Preserve で値を保持）
ReDim Preserve arr(UBound(arr) + 1)
arr(UBound(arr)) = "b"

Dim i As Long
For i = 0 To UBound(arr)
    Debug.Print arr(i)
Next i
```

> `Preserve`を省略すると配列が初期化され、既存の値が消える。値を残したいなら必ず`ReDim Preserve`を使う。

## Hash

```vb
Dim dic as Object
Set dic = CreateObject("Scripting.Dictionary")
'dic.Item(1) = "a"
dic.Add("key", value)

Dim k As Variant
For Each k In dic
    Debug.Print dic.Item(k) 'Debug.Print dic(k)
Next k
```

## Worksheet

宣言1：２行

```vb
Dim ws As Worksheet
Set ws = Sheets("Sheet1")
```

宣言2：１行

```vb
Dim ws As Worksheet: Set ws = Sheets("Sheet1")
```

直接操作１

```vb
Worksheets("Sheet Name")
```

直接操作２

```vb
Sheets("Sheet Name")
```

最終行の取得

```vb
sheet_name = "Sheet Name"
max_row = Worksheets(sheet_name).Cells(Rows.Count, 1).End(xlUp).Row
```

最右列の取得

```vb
Dim lastCol As Long

lastCol = Worksheets(sheet_name).Cells(1, Columns.Count).End(xlToLeft).Column
```

## Workbook

宣言1

```vb
Dim wb as WorkBook
Set wb = Workbooks("Book1.xlsx")
```

宣言2

```vb
Dim wb As Workbook: Set wb = Workbooks("Book1.xlsx")
```

### workbookと、worksheet指定のルール

```vb
Dim wb As Workbook
Dim ws As Worksheet
Set ws = wb.Sheets("SheetName")
' wsには、Sheets("SheetName")が代入されるわけではなく、
' wsには、wb.Sheets("SheetName")が代入されている。
' wb.wsとは書けない。
Debug.Print (ws.Cells(1,1).Value)
'または
Debug.Print (wb.Worksheets(ws.Name).Cells(1,1).Value)
```

> wb.ws と書けないのは、ws は変数名であって、Workbookオブジェクトのメンバーではないからです。Workbookオブジェクトが持つのは Sheets や Worksheets といったコレクションやプロパティであり、任意に作った変数 ws は単なる参照先のラベルにすぎません。

### ファイルを開く

```vb
Dim filename As String
filename = Application.GetOpenFilename(FileFilter:="Excelファイル,*.xls*,CSVファイル,*.csv")
If filename = "False" Then Exit Sub ' キャンセル対応
Dim wb As Workbook
Set wb = Workbooks.Open(filename)
Dim ws As Worksheet
Set ws = wb.Sheets("Sheet1")
' 処理

wb.Close
```

### 特殊参照

- Workbook :this workbook
  - `ThisWorkbook`
- String :location
  - `ThisWorkbook.Path`

## ファイル操作

### ファイルの存在有無

```vb
Dim str As String: str = "filename"
If Dir(str) <> "" Then
    ' 有り
Else
    ' 無し
End If
```

ファイルが無いときのみ実行したい処理は、

```vb
If Dir(str) = "" Then
    ' 無し
End If
```

### ファイル保存するとき

- ドキュメントフォルダ下
  - `"book.xlsx"` or `".\book.xlsx"`
- 実行ファイルと同じ場所
  - `ThisWorkbook.Path & "book.xlsx"`
- 絶対パス
  - `"C:\book.xlsx"`

### 列ファイル書き出し(W)

```vb
Sub OutputColumn()
  Dim column As Long: column = ActiveCell.column
  Dim ws As Worksheet: Set ws = ActiveSheet
  Dim max_row As Long: max_row = ws.Cells(Rows.Count, column).End(xlUp).Row

  Dim output_filename As String
  output_filename = ws.Name + "_" + Format(Time, "hhmmss") + ".txt"
  output_filename = Replace(output_filename, " ", "")

  Open output_filename For Output As #1
  Dim cell_val As String
  For i = 1 To max_row
    cell_val = ws.Cells(i, column).Value
    If Len(cell_val) <> 0 Then
      Print #1, cell_val
    End If
  Next
  Close #1
End Sub
```

## 出力・ダイアログ

コンソール

```vb
Debug.print("a")
```

ダイアログ

```vb
MsgBox("b")
```

## 印刷

```vb
' プリンタのチェック
MsgBox Application.ActivePrinter
' プリンタの指定
Application.ActivePrinter = ""
' sheetの印刷
ActiveSheet.PrintOut
```

## キーボード操作

```vb
Sub Saiban()
  For i = 12 To 34
    SendKeys (i)
    SendKeys ("{DOWN}")
  Next i
End Sub
```

## シート一括操作

シート名全部に対して、名前置換`before`->`after`

```vb
Dim i As Long
For i = 1 To Worksheets.Count
  sn = Worksheets(i).Name
  Worksheets(i).Name = WorksheetFunction.Substitute(sn, "before", "after")
Next
```

コンソールに、シート名一覧する。

```vb
Dim i As Long
For i = 1 To Worksheets.Count
    Debug.Print Worksheets(i).Name
Next
```
