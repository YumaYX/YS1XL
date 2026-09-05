# autogui

- 目的: Win32 API（user32）を使ってマウスのカーソル位置設定・移動・クリックを操作するための宣言。

## 入力

なし。

## 出力

- 型: なし（宣言のみ）
- 内容: マウス操作に使う関数・型・定数を提供します。

## 使用例

```vba
' カーソルを(100, 100)に移動
SetCursorPos 100, 100

' 左クリックを実行
mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
```
