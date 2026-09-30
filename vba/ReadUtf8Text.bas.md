# ReadUtf8Text

- 目的: UTF-8 エンコードのテキストファイルを読み込み、その内容を返す。

## 入力

| 引数名 | 型 | 説明 |
|--------|----|------|
| filePath | String | 読み込むファイルのフルパス |

## 出力

- 型: String
- 内容: ファイルのテキスト内容。
- ファイルが存在しない場合、および読み取り中にエラーが発生した場合は空文字 `""` を返す
  （ストリームは確実に解放される）。

## 依存

- `ADODB.Stream`（`CreateObject("ADODB.Stream")` による遅延バインド。
  「Microsoft ActiveX Data Objects」の参照設定は不要だが、MDAC / ACE が
  Windows にインストールされている必要がある）

## 使用例

```vba
Dim t As String
t = ReadUtf8Text("C:\Temp\data.txt")
Debug.Print t
```
