# IPaddress

IPv4 アドレスとサブネットマスクの検証・相互変換を行う関数群。

含まれる関数: `IsValidIPAddress` / `IsValidSubnetMask` / `IsValidNetworkAddress` / `CIDR2Mask` / `Mask2CIDR`

---

## IsValidIPAddress

- 目的: 引数が有効な IPv4 アドレスか判定する。

### 入力

| 引数名 | 型 | 説明 |
|--------|----|------|
| ip | String | 検証する IP アドレス文字列 |

### 出力

- 型: Boolean
- 内容: 有効な IPv4 アドレスなら True、そうでなければ False。
- 桁数が 4 以外、数値でない、0〜255 の範囲外、先頭ゼロ付き（例: `01.2.3.4`）は False。

### 使用例

```vba
Debug.Print IsValidIPAddress("192.168.0.1")  ' True
Debug.Print IsValidIPAddress("01.2.3.4")     ' False
```

---

## IsValidSubnetMask

- 目的: 引数が有効なサブネットマスクか判定する（上位ビットからの連続した 1 のみ許可）。

### 入力

| 引数名 | 型 | 説明 |
|--------|----|------|
| mask | String | 検証するサブネットマスク文字列 |

### 出力

- 型: Boolean
- 内容: 有効なサブネットマスクなら True、そうでなければ False。
- `0.0.0.0`（`/0`）と `255.255.255.255`（`/32`）も有効として扱う。
  `CIDR2Mask` の出力範囲と整合させるための仕様。
- 1 の途中に 0 が現れる（例: `255.0.255.0`）場合は False。

### 使用例

```vba
Debug.Print IsValidSubnetMask("255.255.255.0")  ' True
Debug.Print IsValidSubnetMask("255.0.255.0")     ' False
Debug.Print IsValidSubnetMask("0.0.0.0")         ' True（/0）
Debug.Print IsValidSubnetMask("255.255.255.255") ' True（/32）
```

---

## IsValidNetworkAddress

- 目的: IP アドレスが指定サブネットマスクのネットワークアドレスか判定する。
- 依存: `IsValidIPAddress` / `IsValidSubnetMask` を使用します。

### 入力

| 引数名 | 型 | 説明 |
|--------|----|------|
| ip | String | 検証する IP アドレス |
| mask | String | サブネットマスク |

### 出力

- 型: Boolean
- 内容: ネットワークアドレス（ホストビットがすべて 0）なら True、そうでなければ False。

> 注意: `192.168.0.1` は `255.255.255.0` のネットワークアドレスではありません。
> ネットワークアドレスは `192.168.0.0` です。

### 使用例

```vba
Debug.Print IsValidNetworkAddress("192.168.0.0", "255.255.255.0")  ' True
Debug.Print IsValidNetworkAddress("192.168.0.1", "255.255.255.0")  ' False
```

---

## CIDR2Mask

- 目的: CIDR 表記（接頭辞長）をサブネットマスク文字列に変換する。

### 入力

| 引数名 | 型 | 説明 |
|--------|----|------|
| cidr | Integer | 接頭辞長（0〜32） |

### 出力

- 型: String
- 内容: ドット区切りのサブネットマスク（例: 24 → `"255.255.255.0"`）。
- 範囲外（0 未満・32 超）の場合は空文字 `""` を返す。

### 使用例

```vba
Debug.Print CIDR2Mask(24)  ' 255.255.255.0
Debug.Print CIDR2Mask(8)   ' 255.0.0.0
Debug.Print CIDR2Mask(0)   ' 0.0.0.0
Debug.Print CIDR2Mask(32)  ' 255.255.255.255
Debug.Print CIDR2Mask(33)  ' （空文字）
```

---

## Mask2CIDR

- 目的: サブネットマスクを CIDR 表記（接頭辞長）に変換する。
- 依存: `IsValidSubnetMask` でマスクの妥当性を検証します。

### 入力

| 引数名 | 型 | 説明 |
|--------|----|------|
| mask | String | サブネットマスク文字列 |

### 出力

- 型: Integer
- 内容: マスク中の 1 ビット数（例: `"255.255.255.0"` → 24）。
- 形式不正、または 1 の連続でない場合は `-1` を返す。

### 使用例

```vba
Debug.Print Mask2CIDR("255.255.255.0")  ' 24
Debug.Print Mask2CIDR("255.0.255.0")    ' -1
Debug.Print Mask2CIDR("not-a-mask")     ' -1
```

---

## 内部ヘルパー（Private）

| 関数 | 型 | 説明 |
|------|----|------|
| `Dec2Bin8` | String | 0〜255 を 8 桁の 2 進文字列に変換する。`IsValidSubnetMask` が内部で使用する。 |

> 注: 従来は `WorksheetFunction.Dec2Bin` を使用していたが、
> Excel の分析ツール（Analysis ToolPak）依存になるため
> 自身のビット演算のみで実装している。
