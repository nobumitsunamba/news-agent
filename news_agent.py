#!/usr/bin/env python3
"""
生成AI日報エージェント
毎朝、生成AI関連ニュースをRSSフィードから収集し、
Claude APIで要約してGmailで送信する。
"""

import os
import smtplib
import logging
from datetime import datetime, timedelta, timezone
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import anthropic
import feedparser
from dotenv import load_dotenv

# ロギング設定
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger(__name__)

# .envファイルの読み込み（GitHub Actionsでは環境変数として直接設定される）
load_dotenv()

# ──────────────────────────────────────────────
# 設定
# ──────────────────────────────────────────────
RSS_FEEDS = [
    "https://news.google.com/rss/search?q=%E7%94%9F%E6%88%90AI&hl=ja&gl=JP&ceid=JP:ja",
    "https://news.google.com/rss/search?q=ChatGPT+OR+Claude+OR+Gemini+OR+LLM&hl=ja&gl=JP&ceid=JP:ja",
    "https://rss.itmedia.co.jp/rss/2.0/news_bursts.xml",
    "https://jp.techcrunch.com/feed/",
]

# 生成AI関連キーワード（いずれかが含まれれば対象）
AI_KEYWORDS = [
    "生成AI",
    "ChatGPT",
    "Claude",
    "Gemini",
    "LLM",
    "大規模言語モデル",
    "Anthropic",
    "OpenAI",
    "Google DeepMind",
    "GPT",
    "AI",                   # 広めにとっておき Claude で絞り込む
    "人工知能",
    "機械学習",
    "ディープラーニング",
    "深層学習",
    "Grok",
    "Llama",
    "Mistral",
    "Copilot",
    "Sora",
    "Stable Diffusion",
    "画像生成",
    "音声生成",
    "自然言語処理",
]

CLAUDE_MODEL = "claude-sonnet-4-6"

# 日本標準時
JST = timezone(timedelta(hours=9))


# ──────────────────────────────────────────────
# ニュース収集
# ──────────────────────────────────────────────
def fetch_articles_from_feed(url: str, since: datetime) -> list[dict]:
    """RSSフィードから記事を取得し、since以降のものを返す。"""
    try:
        feed = feedparser.parse(url)
    except Exception as e:
        logger.warning(f"フィード取得失敗: {url} — {e}")
        return []

    articles = []
    for entry in feed.entries:
        # 公開日時の取得
        pub = entry.get("published_parsed") or entry.get("updated_parsed")
        if pub:
            pub_dt = datetime(*pub[:6], tzinfo=timezone.utc).astimezone(JST)
            if pub_dt < since:
                continue  # 対象日より古い記事はスキップ
        # タイトルとURLの取得
        title = entry.get("title", "").strip()
        link = entry.get("link", "").strip()
        summary = entry.get("summary", "").strip()
        if not title or not link:
            continue
        articles.append({"title": title, "url": link, "summary": summary})

    logger.info(f"  {len(articles)} 件取得: {url[:60]}...")
    return articles


def collect_all_articles() -> list[dict]:
    """全フィードから前日分の記事を収集する。"""
    now_jst = datetime.now(JST)
    # 前日0:00 〜 当日0:00（JST）
    today_midnight = now_jst.replace(hour=0, minute=0, second=0, microsecond=0)
    yesterday_midnight = today_midnight - timedelta(days=1)

    logger.info(f"収集期間: {yesterday_midnight} 〜 {today_midnight} (JST)")

    all_articles: list[dict] = []
    for url in RSS_FEEDS:
        articles = fetch_articles_from_feed(url, yesterday_midnight)
        all_articles.extend(articles)

    logger.info(f"合計 {len(all_articles)} 件の記事を収集")
    return all_articles


# ──────────────────────────────────────────────
# キーワードフィルタリング
# ──────────────────────────────────────────────
def is_ai_related(article: dict) -> bool:
    """タイトルまたはサマリにAIキーワードが含まれるか判定する。"""
    text = (article["title"] + " " + article["summary"]).lower()
    return any(kw.lower() in text for kw in AI_KEYWORDS)


def filter_articles(articles: list[dict]) -> list[dict]:
    """生成AI関連記事のみに絞り込む。"""
    filtered = [a for a in articles if is_ai_related(a)]
    logger.info(f"フィルタリング後: {len(filtered)} 件")
    return filtered


# ──────────────────────────────────────────────
# Claude APIによる要約・整理
# ──────────────────────────────────────────────
def summarize_with_claude(articles: list[dict]) -> str:
    """Claude APIを使って重複排除・重要度ソート・要約を行う。"""
    if not articles:
        return "本日の生成AI関連ニュースはありませんでした。"

    client = anthropic.Anthropic(api_key=os.environ["ANTHROPIC_API_KEY"])

    # Claude への入力用にニュースリストを整形
    news_lines = []
    for i, a in enumerate(articles, 1):
        news_lines.append(f"{i}. タイトル: {a['title']}\n   URL: {a['url']}")

    news_text = "\n\n".join(news_lines)

    prompt = f"""あなたは生成AI専門のニュースキュレーターです。
以下のニュース記事リストを処理してください。

【処理内容】
1. 重複または内容が非常に似ている記事を除外し、代表的な1件のみ残す
2. 重要度・影響度が高い順に並び替える
3. 各記事について3行以内の日本語要約を作成する
4. 生成AIに関係しない記事があれば除外する

【出力形式】
以下の形式で各ニュースを出力してください。番号は1から振り直してください。

---
## [番号]. [タイトル]

[3行以内の要約]

🔗 [URL]
---

【ニュース記事リスト】
{news_text}

それでは処理を開始してください。"""

    logger.info("Claude APIに要約を依頼中...")
    message = client.messages.create(
        model=CLAUDE_MODEL,
        max_tokens=4096,
        messages=[{"role": "user", "content": prompt}],
    )
    result = message.content[0].text
    logger.info("Claude APIから要約を取得しました")
    return result


# ──────────────────────────────────────────────
# メール作成・送信
# ──────────────────────────────────────────────
def build_email_body(summary: str, target_date: datetime) -> tuple[str, str]:
    """メールの件名と本文を生成する。"""
    date_str = target_date.strftime("%Y年%m月%d日")
    subject = f"【生成AI日報】{date_str}"

    body = f"""━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
🤖 生成AI日報 — {date_str}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

{summary}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
このメールは news-agent によって自動送信されました。
https://github.com/nobumitsunamba/news-agent
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
"""
    return subject, body


def send_email(subject: str, body: str) -> None:
    """GmailのSMTPでメールを送信する。"""
    gmail_address = os.environ["GMAIL_ADDRESS"]
    gmail_app_password = os.environ["GMAIL_APP_PASSWORD"]
    to_email = os.environ["TO_EMAIL"]

    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = gmail_address
    msg["To"] = to_email

    # プレーンテキスト
    msg.attach(MIMEText(body, "plain", "utf-8"))

    # HTMLバージョン（Markdownを簡易的にHTML化）
    html_body = body.replace("\n", "<br>")
    html_body = f"<html><body><pre style='font-family:sans-serif;white-space:pre-wrap'>{html_body}</pre></body></html>"
    msg.attach(MIMEText(html_body, "html", "utf-8"))

    logger.info(f"メール送信中: {to_email}")
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(gmail_address, gmail_app_password)
        server.sendmail(gmail_address, to_email, msg.as_string())
    logger.info("メール送信完了")


# ──────────────────────────────────────────────
# メイン処理
# ──────────────────────────────────────────────
def main() -> None:
    logger.info("=== 生成AI日報エージェント 開始 ===")

    # 必須環境変数チェック
    required_vars = ["ANTHROPIC_API_KEY", "GMAIL_ADDRESS", "GMAIL_APP_PASSWORD", "TO_EMAIL"]
    missing = [v for v in required_vars if not os.environ.get(v)]
    if missing:
        raise EnvironmentError(f"環境変数が設定されていません: {', '.join(missing)}")

    # 1. ニュース収集
    articles = collect_all_articles()

    # 2. キーワードフィルタリング
    filtered = filter_articles(articles)

    # 3. Claude APIで要約・整理
    summary = summarize_with_claude(filtered)

    # 4. メール送信
    yesterday_jst = (datetime.now(JST) - timedelta(days=1)).replace(
        hour=0, minute=0, second=0, microsecond=0
    )
    subject, body = build_email_body(summary, yesterday_jst)
    send_email(subject, body)

    logger.info("=== 生成AI日報エージェント 完了 ===")


if __name__ == "__main__":
    main()
