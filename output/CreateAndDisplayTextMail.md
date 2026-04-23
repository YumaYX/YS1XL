# 関数名: CreateAndDisplayTextMail
## 概要
Outlookアプリケーションを利用して、指定された宛先、件名、本文を持つ新しいメールアイテムを作成し、画面上に表示します。本関数は、Outlookオブジェクトへのアクセスとメールプロパティの設定を行います。

## 構文
```vba
Sub CreateAndDisplayTextMail(toAddr As String, _
                             Optional ccAddr As String = "", _
                             Optional bccAddr As String = "", _
                             Optional subjTxt As String = "", _
                             Optional bodyTxt As String = "")
```

## 引数 (Parameters)
| 引数名 | 型 | 必須/任意 | 説明 |
| :--- | :--- | :--- | :--- |
| `toAddr` | `String` | 必須 | 宛先アドレス。カンマ区切りでの指定が可能です。 |
| `ccAddr` | `String` | 任意 | CC（カーボンコピー）アドレス。省略可能です。 |
| `bccAddr` | `String` | 任意 | BCC（ブラインドカーボンコピー）アドレス。省略可能です。 |
| `subjTxt` | `String` | 任意 | メール件名（タイトル）。省略可能です。 |
| `bodyTxt` | `String` | 任意 | メール本文。省略可能です。 |

## 戻り値 (Return Value)
なし (Subプロシージャのため、戻り値はありません)。

## 処理フロー
1.  Outlookアプリケーションオブジェクトを取得または生成します。
2.  新しいメールアイテム（`olMailItem`）を作成します。
3.  引数に基づき、メールのTo、CC、BCC、件名、本文を設定します。
4.  本文の形式をプレーンテキスト（`olFormatPlain`）に設定します。
5.  作成したメールアイテムを画面に表示します（`.Display`）。

## 備考
*   この関数は、実行環境にMicrosoft Outlookがインストールされている必要があります。
*   `On Error Resume Next`が使用されているため、Outlookオブジェクトの取得失敗時など、一部のエラーはスキップされます。

