# ControlFile

ファイルまたはフォルダのコピー・移動・削除を行う関数群。
パスの種類（ファイル／フォルダ）は自動で判定する。

含まれる関数: `Cp` / `Mv` / `Rm`

---

## Cp

- 目的: ファイルまたはフォルダをコピーする（パスの種類を自動判定）。

### 入力

| 引数名 | 型 | 説明 |
|--------|----|------|
| sourcePath | String | コピー元のパス |
| destinationPath | String | コピー先のパス |
| overwrite（省略可） | Boolean | 既存ファイル/フォルダの上書き可否（省略時は False） |

### 出力

- 型: Boolean
- 内容: 成功で True、失敗で False。
- コピー元が存在しない場合やコピー先に同名がある場合の失敗は False になる。

### 使用例

```vba
If Cp("C:\tmp\a.txt", "C:\tmp\b.txt", True) Then
    Debug.Print "コピー成功"
End If
```

---

## Mv

- 目的: ファイルまたはフォルダを移動する（パスの種類を自動判定）。

### 入力

| 引数名 | 型 | 説明 |
|--------|----|------|
| sourcePath | String | 移動元のパス |
| destinationPath | String | 移動先のパス |

### 出力

- 型: Boolean
- 内容: 成功で True、失敗で False。

### 使用例

```vba
If Mv("C:\tmp\a.txt", "C:\tmp\b.txt") Then
    Debug.Print "移動成功"
End If
```

---

## Rm

- 目的: ファイルまたはフォルダを削除する（パスの種類を自動判定）。

### 入力

| 引数名 | 型 | 説明 |
|--------|----|------|
| targetPath | String | 削除対象のパス |

### 出力

- 型: Boolean
- 内容: 成功で True、失敗で False。

### 使用例

```vba
If Rm("C:\tmp\a.txt") Then
    Debug.Print "削除成功"
End If
```

> 注: フォルダを削除する場合は中身もまとめて削除する（`DeleteFolder` の第2引数に `True` を渡す）。
