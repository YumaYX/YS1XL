---
layout: default
permalink: /mechanism/
---

# 概要 - YS1XL

本プロジェクトはVBAの関数集である。Microsoft Excel向けに設計している。すぐに使える再利用可能なコンポーネントを提供する。それにより、同じコードを繰り返し記述する手間を省く。これらの関数を活用することで、定型作業の自動化やワークフローの効率化が可能になる。また、一貫性と保守性の向上にも寄与し、Excelベースの業務をより効率的かつ信頼性の高いものにする。

<https://yumayx.github.io/YS1XL/>

# 生成の方針

## AIエージェントOpenCode

OpenCodeは、オープンソースのAIコーディングエージェントである。ユーザーのコード作成を支援する。無料モデルが提供されている。また、任意のプロバイダーの任意のモデルに接続でき、Claude、GPT、Geminiなどにも対応できる。

OpenCodeはコードやコンテキストデータを保存しないことになっている。
しかし、使用するモデルによっては、データ収集や学習が行なわれているケースがあるため使用には、注意が必要である。

<https://opencode.ai/ja>

本プロジェクトでは、AIエージェントOpenCodeを使う。VBAのプログラム、そのプログラムの説明・使用例は、AIエージェントを使用して生成される。

### 使用するMCPサーバー

MCPサーバーは`Filesystem MCP Server`を使用している。AIエージェントとファイルシステム間のインターフェースとして、ファイル操作をスムーズに行うために導入する。

以下は、導入の際のコマンドである。実行環境は`AlmaLinux 10`である。

```sh
dnf -y update
dnf install nodejs npm -y
npm install -g opencode-ai
npm install -g @modelcontextprotocol/server-filesystem
```

#### ~/.config/opencode/opencode.json

`/work` は AIエージェントに許可する作業ディレクトリ。環境に合わせて変更すること。

```json
{
  "mcp": {
    "filesystem": {
      "type": "local",
      "command": [
        "npx",
        "-y",
        "@modelcontextprotocol/server-filesystem",
        "/work"
      ],
      "enabled": true
    }
  }
}
```
