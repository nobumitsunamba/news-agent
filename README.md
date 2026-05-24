# 🤖 生成AI日報エージェント

毎朝 **日本時間 7:00** に生成AI関連ニュースを自動収集・要約してGmailで送信するPythonスクリプトです。

## 機能

- 複数のRSSフィードから前日分の生成AI関連ニュースを収集
- キーワードフィルタリングで無関係な記事を除外
- **Claude API（claude-sonnet-4-20250514）** で重複排除・重要度ソート・日本語3行要約
- GmailのSMTPで自分宛にメール送信
- GitHub Actionsで毎朝自動実行

## ディレクトリ構成

```
news-agent/
├── news_agent.py                    # メインスクリプト
├── requirements.txt                 # Pythonパッケージ
├── .github/
│   └── workflows/
│       └── news_agent.yml           # GitHub Actions ワークフロー
├── .gitignore
└── README.md
```

## セットアップ手順

### 1. リポジトリをクローン

```bash
git clone https://github.com/nobumitsunamba/news-agent.git
cd news-agent
```

### 2. Python パッケージをインストール

```bash
pip install -r requirements.txt
```

### 3. `.env` ファイルを作成（ローカル実行用）

`.env` ファイルの内容:

```
ANTHROPIC_API_KEY=sk-ant-xxxxxxxxxxxxxxxxxx
GMAIL_ADDRESS=your-address@gmail.com
GMAIL_APP_PASSWORD=xxxx xxxx xxxx xxxx
TO_EMAIL=recipient@example.com
```

> ⚠️ `.env` は `.gitignore` に含まれているため、GitHubにはアップロードされません。

### 4. Gmail アプリパスワードの取得

Gmailの「2段階認証」が有効な場合、通常のパスワードではSMTP認証できません。
**アプリパスワード**を取得してください。

1. [Googleアカウント管理](https://myaccount.google.com/) にアクセス
2. **セキュリティ** → **2段階認証プロセス** を有効化
3. **セキュリティ** → **アプリパスワード** → 「メール」「Windowsパソコン」などを選択
4. 生成された **16桁のパスワード**（スペース区切り）をコピー
5. `.env` の `GMAIL_APP_PASSWORD` に設定

### 5. ローカルでの動作確認

```bash
python news_agent.py
```

---

## GitHub Secrets の設定（自動実行に必要）

GitHub Actionsでの自動実行には、リポジトリの **Secrets** に環境変数を登録する必要があります。

### 手順

1. GitHubリポジトリの **Settings** タブを開く
2. 左サイドバーの **Secrets and variables** → **Actions** をクリック
3. **New repository secret** ボタンをクリック
4. 以下の4つのSecretを登録する

| Secret 名 | 値の例 | 説明 |
|---|---|---|
| `ANTHROPIC_API_KEY` | `sk-ant-xxxxxxxx` | [Anthropic Console](https://console.anthropic.com/) で取得 |
| `GMAIL_ADDRESS` | `your@gmail.com` | 送信元のGmailアドレス |
| `GMAIL_APP_PASSWORD` | `xxxx xxxx xxxx xxxx` | Gmailのアプリパスワード（16桁） |
| `TO_EMAIL` | `recipient@example.com` | 送信先メールアドレス |

### Secret の登録場所

```
GitHub リポジトリ → Settings → Secrets and variables → Actions → New repository secret
```

---

## GitHub Actions ワークフロー

`.github/workflows/news_agent.yml` で定義されています。

- **自動実行**: 毎朝 UTC 22:00（日本時間 7:00）
- **手動実行**: GitHubの Actions タブ → `生成AI日報` → **Run workflow**

```yaml
on:
  schedule:
    - cron: '0 22 * * *'   # JST 07:00
  workflow_dispatch:        # 手動実行
```

---

## 収集対象 RSSフィード

| フィード | 説明 |
|---|---|
| Google News（生成AI） | 生成AI関連の最新ニュース |
| Google News（ChatGPT/Claude/Gemini/LLM） | 主要AIサービス関連ニュース |
| ITmedia NEWS | IT系メディア |
| TechCrunch Japan | スタートアップ・テック系メディア |

## キーワードフィルタリング

以下のキーワードを含む記事のみが収集対象になります:

`生成AI`, `ChatGPT`, `Claude`, `Gemini`, `LLM`, `大規模言語モデル`,
`Anthropic`, `OpenAI`, `Google DeepMind`, `GPT`, `AI`, `人工知能`,
`機械学習`, `ディープラーニング`, `深層学習`, `Grok`, `Llama`, `Mistral`,
`Copilot`, `Sora`, `Stable Diffusion`, `画像生成`, `音声生成`, `自然言語処理`

---

## メール送信例

```
件名: 【生成AI日報】2025年05月24日

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
🤖 生成AI日報 — 2025年05月24日
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

## 1. OpenAIが新モデル「GPT-5」を発表

OpenAIは最新の大規模言語モデル「GPT-5」を正式発表した。
従来モデルと比較して推論能力が大幅に向上しており、
数学や科学分野でのベンチマークで最高スコアを記録。

🔗 https://example.com/news/...

...
```

---

## トラブルシューティング

### メールが届かない

- `GMAIL_APP_PASSWORD` が正しく設定されているか確認
- Gmailの2段階認証が有効になっているか確認
- 「安全性の低いアプリのアクセス」ではなく **アプリパスワード** を使用しているか確認

### GitHub Actions が実行されない

- Secrets が正しく登録されているか確認（名前の大文字小文字に注意）
- Actions タブで Workflow が有効になっているか確認
- 手動実行（`workflow_dispatch`）で動作確認

### RSS フィードが取得できない

- ネットワーク接続を確認
- フィードURLが有効か確認（ブラウザで直接アクセス）

---

## ライセンス

MIT License