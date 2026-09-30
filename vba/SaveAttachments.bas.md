# SaveAttachments

- 目的: 指定した Outlook メールフォルダ内のメールに添付された、キーワードを含むファイルを保存する。

## 入力

| 引数名 | 型 | 説明 |
|--------|----|------|
| mailFolderPath | String | メールフォルダ（ルートからのパス。`\` 区切り） |
| targetFolderPath | String | 保存先フォルダ（末尾の `\` は除去して扱う） |
| keyword | String | 添付ファイル名のキーワード（大文字小文字を区別しない） |

## 出力

- 型: なし（Sub）
- 内容: 条件に一致する添付ファイルを保存先へ書き出す（同名既存ファイルはスキップ）。
- 何もせず終了する条件:

  - `keyword` が空文字（`InStr` が常に 1 となり全添付が対象になるため）
  - 保存先フォルダが存在しない
  - Outlook が利用できない
  - 指定したメールフォルダが存在しない

## 依存

- Outlook（`Outlook.Application` を `GetObject` → `CreateObject` の順で取得）
- `Scripting.FileSystemObject`（フォルダ存在確認・保存先の重複判定に使用）

## 使用例

```vba
Call SaveAttachments("アーカイブ", "C:\Temp", "キーワード")
```
