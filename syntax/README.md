# Excel VBA チートシート

## find

見つかったセル（Rangeオブジェクト）を返す。

```vb
Set cell = Cells.Find(What:="キー見出し", LookAt:=xlWhole)
val = cell.Offset(0, 1).value
```

```vb
Set cell = Cells.Find(What:="キー見出し", LookAt:=xlPart)
val = cell.Offset(0, 1).value
```

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
Worksheet.Cells.Clear
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
End If
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

str = "hello"
Call f2(str)
```

### 返り値ありパターン

`As 型`をつけ、最後に**関数名へ代入**する。`Set` が必要なのはオブジェクトを返す場合。
`Sub` は値を返せないため、必ず `Function` を使う。

```vb
Function f1() As String
  f1 = "hello"
End Function

Debug.Print f1()
```

オブジェクトを返す場合:

```vb
Function f2() As Object
  Set f2 = CreateObject("Scripting.Dictionary")
End Function

Dim d As Object
Set d = f2()
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
Dim dic As Object
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

宣言1：２行

```vb
Dim wb As Workbook
Set wb = Workbooks("Book1.xlsx")
```

宣言2：１行

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

`GetOpenFilename` はキャンセル時に `Boolean` の `False` を返す。
`String` 型の変数に代入すると `"False"` という文字列になるため、
`Variant` のまま受け取って `VarType` で判定するのが確実。

```vb
Dim filename As Variant
filename = Application.GetOpenFilename(FileFilter:="Excelファイル,*.xls*,CSVファイル,*.csv")

If VarType(filename) = vbBoolean Then Exit Sub ' キャンセル対応

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
  Dim ws As Worksheet: Set ws = ActiveSheet
  Dim column As Long: column = ws.Cells(1, 1).Column

  Dim maxRow As Long
  maxRow = ws.Cells(Rows.Count, column).End(xlUp).Row

  ' 文字列連結は + ではなく & を使う
  Dim outputFilename As String
  outputFilename = ws.Name & "_" & Format(Time, "hhmmss") & ".txt"
  outputFilename = Replace(outputFilename, " ", "")

  Dim fNum As Integer: fNum = FreeFile
  Open outputFilename For Output As #fNum

  Dim cellVal As Variant
  Dim i As Long
  For i = 1 To maxRow
    cellVal = ws.Cells(i, column).Value
    If Len(cellVal) <> 0 Then
      Print #fNum, cellVal
    End If
  Next i

  Close #fNum
End Sub
```

> 注: 文字列の連結に `+` を使うと、`+` が数値加算の意味も持つため
> 想定外の型変換エラーが発生することがある。必ず `&` を使うこと。
> また `For i = 1 To max_row` の `i` は必ず `Dim` で宣言すること
> （`Option Explicit` 有効時は未宣言変数でエラーになる）。

## 出力・ダイアログ

コンソール

```vb
Debug.Print "a"
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
  Dim i As Long
  For i = 12 To 34
    SendKeys CStr(i)
    SendKeys "{DOWN}"
  Next i
End Sub
```

## シート一括操作

シート名全部に対して、名前置換`before`->`after`

```vb
Dim i As Long
Dim sn As String
For i = 1 To Worksheets.Count
  sn = Worksheets(i).Name
  Worksheets(i).Name = WorksheetFunction.Substitute(sn, "before", "after")
Next i
```

コンソールに、シート名一覧する。

```vb
Dim i As Long
For i = 1 To Worksheets.Count
    Debug.Print Worksheets(i).Name
Next i
```

## セル内の改行で分割

選択中の1セルを、セル内の改行（`Alt + Enter`）ごとに分割し、そのセルの下へ書き込む。
元のセルは変更しない。

```vba
Sub SplitCellByLineBreak()
    Dim src As Range
    Dim lines As Variant
    Dim text As String
    Dim i As Long
    Dim n As Long

    If Not TypeOf Selection Is Range Then
        MsgBox "セルを1つ選択してください。", vbExclamation
        Exit Sub
    End If

    If Selection.CountLarge <> 1 Then
        MsgBox "セルを1つだけ選択してください。", vbExclamation
        Exit Sub
    End If

    Set src = Selection

    If IsError(src.Value) Then Exit Sub
    text = CStr(src.Value)
    If Len(text) = 0 Then Exit Sub

    text = Replace(text, vbCrLf, vbLf)
    text = Replace(text, vbCr, vbLf)
    lines = Split(text, vbLf)

    n = UBound(lines) - LBound(lines) + 1
    If src.Row + n > src.Worksheet.Rows.Count Then
        MsgBox "分割後の行がシートの最終行を超えます。", vbExclamation
        Exit Sub
    End If

    src.Offset(1, 0).Resize(n, 1).ClearContents

    For i = LBound(lines) To UBound(lines)
        src.Offset(i - LBound(lines) + 1, 0).Value = lines(i)
    Next i
End Sub
```

