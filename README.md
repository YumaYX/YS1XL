# YS1XL

Reusable procedures and functions for Excel automation.

- [Documents - https://yumayx.github.io/YS1XL/](https://yumayx.github.io/YS1XL/)
- [Syntax](https://github.com/YumaYX/YS1XL/blob/main/syntax/README.md)

## how to use

```sh
curl -o module.bas.txt https://raw.githubusercontent.com/YumaYX/YS1XL/refs/heads/main/module.bas
```

## for development

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
