# OpenExcel

- 目的: ファイル選択ダイアログで選択した Excel/CSV ファイルを開き Workbook を返す。

# 入力（日本語）

なし

# 出力（日本語）

- 型: Workbook
- 内容: 開いたワークブック。キャンセル時は Nothing。

# 使用例

```vba
Dim wb As Workbook
Set wb = OpenExcel()
```