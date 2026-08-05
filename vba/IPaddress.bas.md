# IsValidIPAddress

- 目的: 引数が有効な IPv4 アドレスか判定する。

# 入力（日本語）

| 引数名 | 型   | 説明                       |
|--------|------|----------------------------|
| ip     | String | 検証する IP アドレス文字列 |

# 出力（日本語）

- 型: Boolean
- 内容: 有効な IPv4 アドレスなら True、そうでなければ False。

# 使用例

```vba
Debug.Print IsValidIPAddress("192.168.0.1")  ' True
Debug.Print IsValidIPAddress("01.2.3.4")     ' False
```

---

# IsValidSubnetMask

- 目的: 引数が有効なサブネットマスクか判定する（先頭からの連続した 1 のみ許可）。

# 入力（日本語）

| 引数名 | 型   | 説明                           |
|--------|------|--------------------------------|
| mask   | String | 検証するサブネットマスク文字列 |

# 出力（日本語）

- 型: Boolean
- 内容: 有効なサブネットマスクなら True、そうでなければ False（全 1 / 全 0 は除外）。

# 使用例

```vba
Debug.Print IsValidSubnetMask("255.255.255.0")  ' True
Debug.Print IsValidSubnetMask("255.0.255.0")     ' False
```

---

# IsValidNetworkAddress

- 目的: IP アドレスが指定サブネットマスクのネットワークアドレスか判定する。

# 入力（日本語）

| 引数名 | 型   | 説明                           |
|--------|------|--------------------------------|
| ip     | String | 検証する IP アドレス            |
| mask   | String | サブネットマスク               |

# 出力（日本語）

- 型: Boolean
- 内容: ネットワークアドレスに該当すれば True、そうでなければ False。

# 使用例

```vba
Debug.Print IsValidNetworkAddress("192.168.0.1", "255.255.255.0")  ' False
Debug.Print IsValidNetworkAddress("192.168.0.0", "255.255.255.0")   ' True
```

---

# CIDR2Mask

- 目的: CIDR 表記（接頭辞長）をサブネットマスク文字列に変換する。

# 入力（日本語）

| 引数名 | 型      | 説明              |
|--------|---------|-------------------|
| cidr   | Integer | 接頭辞長（0〜32）|

# 出力（日本語）

- 型: String
- 内容: ドット区切りのサブネットマスク（例: 24 → "255.255.255.0"）。

# 使用例

```vba
Debug.Print CIDR2Mask(24)  ' 255.255.255.0
```

---

# Mask2CIDR

- 目的: サブネットマスクを CIDR 表記（接頭辞長）に変換する。

# 入力（日本語）

| 引数名 | 型   | 説明                           |
|--------|------|--------------------------------|
| mask   | String | サブネットマスク文字列         |

# 出力（日本語）

- 型: Integer
- 内容: マスク中の連続する 1 ビット数を返す（例: "255.255.255.0" → 24）。

# 使用例

```vba
Debug.Print Mask2CIDR("255.255.255.0")  ' 24
```