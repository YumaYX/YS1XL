# OpenExcel

- 目的: ファイル選択ダイアログで選択した Excel/CSV ファイルを開き Workbook を返す。

## 入力

なし。

## 出力

- 型: Workbook
- 内容: 開いたワークブック。キャンセル時は `Nothing`。

> 注: `Application.GetOpenFilename` はキャンセル時に `Boolean` の `False` を返すため、
> `VarType` で判定している。戻り値がオブジェクト型なので、
> キャンセル時には `Set` で `Nothing` を代入して返す。

## 使用例

```vba
Dim wb As Workbook
Set wb = OpenExcel()
If wb Is Nothing Then Exit Sub   ' キャンセル時
```
