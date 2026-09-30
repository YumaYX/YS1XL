# YS1XL

Reusable procedures and functions for Excel automation.

- [Documents - https://yumayx.github.io/YS1XL/](https://yumayx.github.io/YS1XL/)
- [Syntax](https://github.com/YumaYX/YS1XL/blob/main/syntax/README.md)

## how to use

`module.bas` は `vba/*.bas` を連結した生成物。VBA エディタに貼り付けて取り込む。

```sh
curl -o module.bas https://raw.githubusercontent.com/YumaYX/YS1XL/refs/heads/main/module.bas
```

## for development

`module.bas` は生成物。**編集は `vba/*.bas` に対して行い**、
`make concat` で `module.bas` を再生成する。

### split

module.bas -> vba/*.bas

```sh
make split
```

### concat

vba/*.bas -> module.bas

```sh
make concat
```

> 注意: `make split` は `vba/*.bas` を上書きする。
> `vba/*.bas` に未コミットの変更がある場合は実行しないこと。

### ドキュメント

`_posts/` は `vba/*.bas.md` から生成される（`make posts`）。
`.bas` に対して説明が無いファイルは `make describe` で生成を試みる
（外部ツール `ys-ollama2file` が必要）。

### 改行コード

`.gitattributes` により `*.bas` は CRLF（VBE と同じ）、
マークダウン系は LF で保持する。
