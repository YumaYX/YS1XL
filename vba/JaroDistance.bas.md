# JaroDistance

- 目的: 2つの文字列間のJaro距離（0〜1）を返す。0=完全一致、1=全く異なる。

## 入力

| 引数名 | 型 | 説明 |
|--------|----|------|
| s1 | String | 比較元の文字列 |
| s2 | String | 比較先の文字列 |

## 出力

- 型: Double
- 内容: Jaro距離。0（完全一致）〜1（共通文字なし）。共通文字があれば類似度の補数（1 − Jaro類似度）。
- 両方が空文字列の場合は 0（完全一致扱い）。
- 片方だけが空文字列の場合は共通文字なしなので 1。

## 使用例

```vba
Debug.Print JaroDistance("MARTHA", "MARHTA")  ' ≈ 0.0556（転置を含む例）
Debug.Print JaroDistance("abc", "abc")        ' 0
Debug.Print JaroDistance("abc", "xyz")        ' 1
```