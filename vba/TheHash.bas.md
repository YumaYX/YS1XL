# TheHash

- 目的: ワークシートの指定列をキーとした連想配列（ディクショナリ）を作成する。

## 入力

| 引数名 | 型 | 説明 |
|--------|----|------|
| ws | Worksheet | 対象のワークシート |
| keyIndex | Long | キーとする列番号 |

## 出力

- 型: Object（Scripting.Dictionary）
- 内容: Key=セルの値を文字列化したもの、Item=その値が最後に現れた行番号。
- 同じ値が複数ある場合は、最後に出現した行番号で上書きされる。
- エラー値（`#N/A` など）はキーに登録されず、除外される。
- 空列のときは空の Dictionary を返す。

## 依存

- `Scripting.Dictionary`（`CreateObject` による遅延バインド。参照設定は不要）
- `Application.WorksheetFunction.CountA`（空列判定に使用）

## 使用例

```vba
Dim myHash As Object
Set myHash = TheHash(Sheets("Sheet1"), 1)
Debug.Print myHash("someKey")  ' キーに対応する行番号が返る
```
