# OpenExcel

- 目的: ファイル選択ダイアログを表示し、選択した Excel／CSV ファイルを開いて Workbook を返す。

# 入力（日本語）

| 引数名 | 型 | 説明 |
|--------|----|------|

# 出力（日本語）

- 型: Workbook
- 内容: 開いたワークブック。ユーザーがキャンセルした場合は Nothing を返す。

# 使用例

```vba
Dim wb As Workbook
Set wb = OpenExcel()
```