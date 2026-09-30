# AutoGui

- 目的: Win32 API（user32）を使ってマウスのカーソル位置取得・移動・クリックを操作するための宣言。

## 入力

なし。

## 出力

- 型: なし（宣言のみ。`Sub` / `Function` として呼び出す）
- 内容: 次の Private メンバを提供する。

| メンバ | 種別 | 説明 |
|--------|------|------|
| `SetCursorPos` | Function | カーソル位置をピクセル単位で設定する |
| `GetCursorPos` | Function | 現在のカーソル位置を `POINTAPI` に書き出す |
| `mouse_event` | Sub | マウスイベント（クリック等）を発生させる |
| `POINTAPI` | Type | カーソル座標を保持するユーザー定義型（X / Y） |
| `MOUSEEVENTF_LEFTDOWN` | Const | 左ボタン下楼のイベントフラグ |
| `MOUSEEVENTF_LEFTUP` | Const | 左ボタン上階のイベントフラグ |
| `MOUSEEVENTF_RIGHTDOWN` | Const | 右ボタン下楼のイベントフラグ |
| `MOUSEEVENTF_RIGHTUP` | Const | 右ボタン上階のイベントフラグ |

> 注: すべて `Private` 宣言のため、同じプロジェクト内の別モジュールからは呼び出せない。
> `AutoGui.bas` と同じ標準モジュールにコードを置くか、`Public` に変更すること。
> また `Declare` は Windows 32bit 版（VBA6）向けの `PtrSafe` 指定で、
> 64bit 版では `LongPtr` を使う必要がある。

## 使用例

```vba
' カーソルを(100, 100)に移動
SetCursorPos 100, 100

' 左クリックを実行
mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
```

```vba
' 現在のカーソル位置を取得
Dim pt As POINTAPI
GetCursorPos pt
Debug.Print pt.X & "," & pt.Y
```

