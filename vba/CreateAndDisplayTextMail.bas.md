# CreateAndDisplayTextMail

- 目的: Outlook で新規テキストメールを作成し画面に表示する（返り値なし）。

# 入力

| 引数名      | 型     | 説明                                   |
|-------------|--------|----------------------------------------|
| toAddr      | String | 宛先（カンマ区切りで複数指定可）       |
| ccAddr（省略可） | String | CC（省略時は空）                  |
| bccAddr（省略可）| String | BCC（省略時は空）                 |
| subjTxt（省略可）| String | タイトル（省略時は空）            |
| bodyTxt（省略可）| String | 本文（省略時は空）                |

# 出力

- 型: なし（Sub）
- 内容: Outlook が未起動なら起動し、テキスト形式の新規メールを表示する。

# 使用例

```vba
CreateAndDisplayTextMail "to@example.com", "", "", "件名", "本文"
```