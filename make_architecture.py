"""
生成AI日報エージェント ITアーキテクチャ図（PowerPoint 1枚）
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE_TYPE
import pptx.oxml.ns as nsmap
from lxml import etree

# ── カラーパレット ──────────────────────────────
C_BG        = RGBColor(0xF4, 0xF7, 0xFB)  # スライド背景
C_NAVY      = RGBColor(0x1B, 0x2A, 0x4A)  # タイトル・見出し
C_BLUE      = RGBColor(0x2E, 0x6E, 0xC8)  # RSS / メイン枠
C_TEAL      = RGBColor(0x00, 0x96, 0x88)  # Claude API
C_ORANGE    = RGBColor(0xF5, 0x7C, 0x00)  # Gmail
C_PURPLE    = RGBColor(0x6A, 0x1B, 0x9A)  # GitHub Actions
C_GRAY      = RGBColor(0x78, 0x90, 0xA8)  # Secrets
C_WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
C_LIGHT_BG  = RGBColor(0xE8, 0xF0, 0xFE)  # 薄い青背景
C_ARROW     = RGBColor(0x90, 0xA4, 0xAE)  # 矢印

SLIDE_W = Inches(13.33)
SLIDE_H = Inches(7.5)

prs = Presentation()
prs.slide_width  = SLIDE_W
prs.slide_height = SLIDE_H

slide = prs.slides.add_slide(prs.slide_layouts[6])  # 白紙レイアウト

# ── 背景色 ─────────────────────────────────────
def set_slide_bg(slide, color: RGBColor):
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = color

set_slide_bg(slide, C_BG)

shapes = slide.shapes

# ── ヘルパー関数 ────────────────────────────────
def add_rect(slide, l, t, w, h, fill, text="", font_size=11, bold=False,
             font_color=C_WHITE, line_color=None, line_width=Pt(0),
             align=PP_ALIGN.CENTER, v_anchor=None, radius=None):
    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE_TYPE.RECTANGLE
        Inches(l), Inches(t), Inches(w), Inches(h)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    if line_color:
        shape.line.color.rgb = line_color
        shape.line.width = line_width
    else:
        shape.line.fill.background()

    if radius:
        # 角丸
        sp = shape._element
        prstGeom = sp.find('.//' + nsmap.qn('a:prstGeom'))
        if prstGeom is not None:
            prstGeom.set('prst', 'roundRect')
            avLst = prstGeom.find(nsmap.qn('a:avLst'))
            if avLst is None:
                avLst = etree.SubElement(prstGeom, nsmap.qn('a:avLst'))
            gd = etree.SubElement(avLst, nsmap.qn('a:gd'))
            gd.set('name', 'adj')
            gd.set('fmla', f'val {radius}')

    if text:
        tf = shape.text_frame
        tf.word_wrap = True
        if v_anchor:
            tf.vertical_anchor = v_anchor
        p = tf.paragraphs[0]
        p.alignment = align
        run = p.add_run()
        run.text = text
        run.font.size = Pt(font_size)
        run.font.bold = bold
        run.font.color.rgb = font_color
    return shape

def add_textbox(slide, l, t, w, h, text, font_size=10, bold=False,
                color=C_NAVY, align=PP_ALIGN.CENTER):
    txBox = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = color
    return txBox

def add_arrow(slide, x1, y1, x2, y2, color=C_ARROW, width=Pt(1.5)):
    """始点→終点の矢印ラインを追加"""
    from pptx.util import Inches
    connector = slide.shapes.add_connector(
        1,  # MSO_CONNECTOR_STRAIGHT
        Inches(x1), Inches(y1), Inches(x2), Inches(y2)
    )
    connector.line.color.rgb = color
    connector.line.width = width
    # 矢印先端
    ln = connector.line._ln
    tailEnd = ln.find(nsmap.qn('a:tailEnd'))
    if tailEnd is None:
        tailEnd = etree.SubElement(ln, nsmap.qn('a:tailEnd'))
    tailEnd.set('type', 'none')
    headEnd = ln.find(nsmap.qn('a:headEnd'))
    if headEnd is None:
        headEnd = etree.SubElement(ln, nsmap.qn('a:headEnd'))
    headEnd.set('type', 'arrow')
    headEnd.set('w', 'med')
    headEnd.set('len', 'med')
    return connector

# ══════════════════════════════════════════════
# タイトルバー
# ══════════════════════════════════════════════
add_rect(slide, 0, 0, 13.33, 0.7, fill=C_NAVY,
         text="生成AI日報エージェント  —  ITアーキテクチャ",
         font_size=20, bold=True, font_color=C_WHITE)

# ══════════════════════════════════════════════
# ゾーン背景
# ══════════════════════════════════════════════
# 左ゾーン（入力）薄枠
add_rect(slide, 0.2, 0.85, 3.2, 5.8, fill=RGBColor(0xE3,0xEA,0xF8),
         line_color=RGBColor(0xBB,0xCC,0xEE), line_width=Pt(1))
add_textbox(slide, 0.3, 0.88, 2.5, 0.3, "📡  入力（RSSフィード）",
            font_size=9, bold=True, color=C_BLUE)

# 中央ゾーン（処理）薄枠
add_rect(slide, 3.65, 0.85, 3.8, 5.8, fill=RGBColor(0xE8,0xF5,0xE9),
         line_color=RGBColor(0xBB,0xDD,0xBB), line_width=Pt(1))
add_textbox(slide, 3.75, 0.88, 3.2, 0.3, "⚙️  処理（news_agent.py）",
            font_size=9, bold=True, color=RGBColor(0x2E,0x7D,0x32))

# 右ゾーン（出力）薄枠
add_rect(slide, 9.7, 0.85, 3.4, 5.8, fill=RGBColor(0xFF, 0xF3, 0xE0),
         line_color=RGBColor(0xEE, 0xCC, 0xAA), line_width=Pt(1))
add_textbox(slide, 9.8, 0.88, 3.0, 0.3, "📤  出力",
            font_size=9, bold=True, color=C_ORANGE)

# ══════════════════════════════════════════════
# 左：RSSフィード（4つ）
# ══════════════════════════════════════════════
rss_items = [
    ("Google News\n生成AI", 1.35),
    ("Google News\nChatGPT/Claude\n/Gemini/LLM", 2.25),
    ("ITmedia NEWS", 3.45),
    ("TechCrunch Japan", 4.35),
]

for label, top in rss_items:
    h = 0.65 if "\n" not in label else (0.85 if label.count("\n") == 2 else 0.65)
    add_rect(slide, 0.35, top, 2.9, h, fill=C_BLUE,
             text=label, font_size=9.5, bold=False, font_color=C_WHITE,
             radius=20000)

# ══════════════════════════════════════════════
# 中央：処理ステップ（4ステップ）
# ══════════════════════════════════════════════
steps = [
    ("① RSS収集\nfetch_articles_from_feed()",  1.35, RGBColor(0x43,0xA0,0x47)),
    ("② キーワードフィルタ\nfilter_articles()",        2.35, RGBColor(0x2E,0x7D,0x32)),
    ("③ Claude API 要約\nsummarize_with_claude()",   3.35, RGBColor(0x1B,0x5E,0x20)),
    ("④ メール送信\nsend_email()",                    4.55, RGBColor(0x33,0x69,0x1E)),
]

for label, top, col in steps:
    add_rect(slide, 3.8, top, 3.5, 0.72, fill=col,
             text=label, font_size=9.5, bold=False, font_color=C_WHITE,
             radius=15000)

# ステップ間の下矢印
for top in [2.07, 3.07, 4.07]:
    add_arrow(slide, 5.55, top, 5.55, top + 0.28,
              color=RGBColor(0x2E,0x7D,0x32), width=Pt(2.0))

# ══════════════════════════════════════════════
# 右：出力
# ══════════════════════════════════════════════
# Gmailボックス
add_rect(slide, 9.85, 1.35, 3.1, 1.1, fill=C_ORANGE,
         text="Gmail\n受信トレイ", font_size=13, bold=True,
         font_color=C_WHITE, radius=20000)
add_textbox(slide, 9.85, 2.5, 3.1, 0.55,
            "件名：【生成AI日報】\nYYYY年MM月DD日",
            font_size=8.5, color=RGBColor(0x6D,0x4C,0x41))

# メール仕様
add_rect(slide, 9.85, 3.15, 3.1, 1.5, fill=RGBColor(0xFF,0xE0,0xB2),
         text="• SMTP SSL（Port 465）\n• プレーンテキスト＋HTML\n• 毎朝 JST 07:00 に自動送信",
         font_size=9, bold=False, font_color=RGBColor(0x4E,0x34,0x23),
         align=PP_ALIGN.LEFT, radius=10000)

# ══════════════════════════════════════════════
# 下段：GitHub Actions + Secrets
# ══════════════════════════════════════════════
# GitHub Actions
add_rect(slide, 3.8, 5.75, 3.5, 0.75, fill=C_PURPLE,
         text="⏰  GitHub Actions\ncron: '0 22 * * *' (JST 07:00)",
         font_size=9.5, bold=False, font_color=C_WHITE, radius=15000)

# GitHub Secrets
add_rect(slide, 0.35, 5.75, 2.9, 0.75, fill=C_GRAY,
         text="🔐  GitHub Secrets\nANTHROPIC_API_KEY\nGMAIL_APP_PASSWORD 等",
         font_size=8.5, bold=False, font_color=C_WHITE, radius=15000)

# Claude API
add_rect(slide, 9.85, 5.15, 3.1, 1.35, fill=C_TEAL,
         text="🤖  Claude API\nclaude-sonnet-4-6\n(Anthropic)",
         font_size=10, bold=False, font_color=C_WHITE, radius=20000)

# ══════════════════════════════════════════════
# 矢印
# ══════════════════════════════════════════════
ARROW_COLOR = RGBColor(0x55, 0x77, 0x99)

# RSS → ①RSS収集（各フィードから中央へ）
for src_y in [1.68, 2.58, 3.78, 4.68]:
    add_arrow(slide, 3.25, src_y, 3.8, src_y,
              color=ARROW_COLOR, width=Pt(1.5))

# ④メール送信 → Gmail
add_arrow(slide, 7.3, 4.91, 9.85, 2.6,
          color=C_ORANGE, width=Pt(2.0))

# GitHub Actions → ①（トリガー）
add_arrow(slide, 5.55, 5.75, 5.55, 5.27,
          color=C_PURPLE, width=Pt(2.0))

# GitHub Secrets → 処理ゾーン
add_arrow(slide, 3.25, 6.12, 3.8, 6.12,
          color=C_GRAY, width=Pt(1.5))

# Claude API ↔ ③要約ステップ
add_arrow(slide, 9.85, 5.8, 8.0, 3.71,
          color=C_TEAL, width=Pt(2.0))
add_arrow(slide, 7.3, 3.71, 9.85, 5.6,
          color=C_TEAL, width=Pt(1.5))

# ══════════════════════════════════════════════
# ラベル（矢印の補足）
# ══════════════════════════════════════════════
add_textbox(slide, 7.35, 4.55, 2.4, 0.3,
            "要約・重複排除・\n重要度ソート",
            font_size=7.5, color=C_TEAL)
add_textbox(slide, 7.8, 2.85, 2.0, 0.3,
            "メール本文",
            font_size=7.5, color=C_ORANGE)
add_textbox(slide, 5.6, 5.45, 1.5, 0.28,
            "毎朝トリガー",
            font_size=7.5, color=C_PURPLE)
add_textbox(slide, 3.3, 5.85, 1.3, 0.3,
            "認証情報注入",
            font_size=7.5, color=C_GRAY)

# ══════════════════════════════════════════════
# 凡例（右下）
# ══════════════════════════════════════════════
add_rect(slide, 0.2, 6.9, 13.0, 0.5, fill=C_NAVY,
         text="", font_size=9)
legend_items = [
    ("■ RSSフィード（入力）", C_BLUE, 0.5),
    ("■ 処理（Python）", RGBColor(0x2E,0x7D,0x32), 2.8),
    ("■ Claude API", C_TEAL, 5.0),
    ("■ Gmail（出力）", C_ORANGE, 7.2),
    ("■ GitHub Actions", C_PURPLE, 9.4),
    ("■ GitHub Secrets", C_GRAY, 11.3),
]
for label, color, left in legend_items:
    tb = slide.shapes.add_textbox(Inches(left), Inches(6.97), Inches(2.0), Inches(0.35))
    tf = tb.text_frame
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = label
    run.font.size = Pt(8)
    run.font.color.rgb = color

# ══════════════════════════════════════════════
# 保存
# ══════════════════════════════════════════════
output = "/home/user/news-agent/生成AI日報エージェント_アーキテクチャ.pptx"
prs.save(output)
print(f"✅ 保存完了: {output}")
