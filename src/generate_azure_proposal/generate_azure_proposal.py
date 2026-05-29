from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

OUTPUT_FILE = "Azure導入_提案資料_v1.1.pptx"

# ---------- 基本設定 ----------
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

# 色
NAVY = RGBColor(31, 78, 121)
BLUE = RGBColor(68, 114, 196)
LIGHT_BLUE = RGBColor(221, 235, 247)
DARK_GRAY = RGBColor(64, 64, 64)
GRAY = RGBColor(120, 120, 120)
LIGHT_GRAY = RGBColor(242, 242, 242)
GREEN = RGBColor(0, 176, 80)
ORANGE = RGBColor(237, 125, 49)
RED = RGBColor(192, 0, 0)
WHITE = RGBColor(255, 255, 255)

FONT_JP = "Meiryo"


def set_text_style(run, size=24, bold=False, color=DARK_GRAY, font_name=FONT_JP):
    run.font.name = font_name
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color


def add_title(slide, title, subtitle=None):
    tx = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12.3), Inches(0.7))
    tf = tx.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    r = p.add_run()
    r.text = title
    set_text_style(r, size=28, bold=True, color=NAVY)

    if subtitle:
        tx2 = slide.shapes.add_textbox(
            Inches(0.5), Inches(1.0), Inches(12.3), Inches(0.4)
        )
        tf2 = tx2.text_frame
        tf2.clear()
        p2 = tf2.paragraphs[0]
        r2 = p2.add_run()
        r2.text = subtitle
        set_text_style(r2, size=14, bold=False, color=GRAY)


def add_footer(slide, text="Confidential"):
    tx = slide.shapes.add_textbox(Inches(11.7), Inches(7.0), Inches(1.4), Inches(0.25))
    tf = tx.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.RIGHT
    r = p.add_run()
    r.text = text
    set_text_style(r, size=10, color=GRAY)


def add_bullets(slide, title, bullets, top=1.4, box_height=5.5):
    add_title(slide, title)
    left = Inches(0.8)
    width = Inches(11.8)
    tx = slide.shapes.add_textbox(left, Inches(top), width, Inches(box_height))
    tf = tx.text_frame
    tf.word_wrap = True
    tf.clear()

    for i, b in enumerate(bullets):
        if isinstance(b, tuple):
            text, level = b
        else:
            text, level = b, 0
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.level = level
        p.space_after = Pt(10)
        p.alignment = PP_ALIGN.LEFT
        r = p.add_run()
        r.text = text
        set_text_style(r, size=22 if level == 0 else 19, color=DARK_GRAY)

    add_footer(slide)


def add_label_box(
    slide,
    x,
    y,
    w,
    h,
    text,
    fill,
    line=None,
    font_size=20,
    bold=True,
    color=WHITE,
    radius=True,
):
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE
    box = slide.shapes.add_shape(shape_type, Inches(x), Inches(y), Inches(w), Inches(h))
    box.fill.solid()
    box.fill.fore_color.rgb = fill
    if line is None:
        box.line.color.rgb = fill
    else:
        box.line.color.rgb = line
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    p.space_before = Pt(6)
    p.space_after = Pt(6)
    r = p.add_run()
    r.text = text
    set_text_style(r, size=font_size, bold=bold, color=color)
    return box


def add_small_text(
    slide, x, y, w, h, text, font_size=12, color=GRAY, align=PP_ALIGN.LEFT
):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    set_text_style(r, size=font_size, color=color, bold=False)
    return tb


def add_arrow(slide, x, y, w, h, fill=BLUE):
    arrow = slide.shapes.add_shape(
        MSO_SHAPE.RIGHT_ARROW, Inches(x), Inches(y), Inches(w), Inches(h)
    )
    arrow.fill.solid()
    arrow.fill.fore_color.rgb = fill
    arrow.line.color.rgb = fill
    return arrow


# ---------- Slide 1: 表紙 ----------
slide = prs.slides.add_slide(prs.slide_layouts[6])
band = slide.shapes.add_shape(
    MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(1.2)
)
band.fill.solid()
band.fill.fore_color.rgb = NAVY
band.line.color.rgb = NAVY

tb = slide.shapes.add_textbox(Inches(0.7), Inches(1.7), Inches(12), Inches(1.3))
tf = tb.text_frame
tf.clear()
p = tf.paragraphs[0]
r = p.add_run()
r.text = "Microsoft 365 と連携した Azure 移行のご提案"
set_text_style(r, size=28, bold=True, color=NAVY)

tb2 = slide.shapes.add_textbox(Inches(0.7), Inches(2.9), Inches(12), Inches(1.0))
tf2 = tb2.text_frame
tf2.clear()
p2 = tf2.paragraphs[0]
r2 = p2.add_run()
r2.text = "運用の簡素化・セキュリティ強化・リリース高速化"
set_text_style(r2, size=18, bold=False, color=GRAY)

add_label_box(
    slide,
    8.5,
    1.8,
    3.8,
    0.8,
    "社員のログインを一本化",
    LIGHT_BLUE,
    BLUE,
    18,
    True,
    NAVY,
)
add_label_box(
    slide, 8.5, 2.8, 3.8, 0.8, "鍵をアプリに置かない", LIGHT_BLUE, BLUE, 18, True, NAVY
)
add_label_box(
    slide, 8.5, 3.8, 3.8, 0.8, "安全に自動でリリース", LIGHT_BLUE, BLUE, 18, True, NAVY
)

add_small_text(slide, 0.7, 6.4, 6.5, 0.3, "作成日: 2026-05-29", font_size=11)
add_small_text(slide, 0.7, 6.75, 6.5, 0.3, "作成者: ", font_size=11)
add_footer(slide, "Confidential")

# ---------- Slide 2: 要点 ----------
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_bullets(
    slide,
    "要点（経営判断のための結論）",
    [
        "Microsoft 365 と同じ土台に揃えることで、運用がシンプルになります。",
        "アプリごとにパスワードを持たせず、自動で管理できるため、漏えいリスクを下げられます。",
        "開発から本番までの流れを自動化できるため、リリースが速くなり、人為的ミスも減ります。",
    ],
    top=1.6,
    box_height=4.8,
)

add_label_box(
    slide,
    0.9,
    6.1,
    11.6,
    0.7,
    "結果として、安全性を高めながら、手間と費用の削減が見込めます。",
    LIGHT_BLUE,
    BLUE,
    18,
    True,
    NAVY,
)
add_footer(slide)

# ---------- Slide 3: 現状の課題 ----------
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_bullets(
    slide,
    "現状の課題",
    [
        "システムが複数の環境に分かれているため、管理が複雑になっています。",
        "機密情報が各所に点在し、漏えいリスクと運用コストが高くなっています。",
        "開発から本番への切り替えに手作業が多く、時間がかかります。",
    ],
    top=1.6,
    box_height=4.8,
)
add_label_box(
    slide,
    0.9,
    6.0,
    11.6,
    0.8,
    "分散しているほど、管理の手間とリスクは増えます。",
    RGBColor(255, 242, 204),
    ORANGE,
    18,
    True,
    DARK_GRAY,
)
add_footer(slide)

# ---------- Slide 4: 提案の概要 ----------
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_bullets(
    slide,
    "提案の概要",
    [
        "Microsoft 365 と同じ Azure に移すことで、社員のログインと権限を一本化します。",
        "鍵や証明書を専用の場所で自動管理し、アプリに置かないようにします。",
        "開発から本番への流れを自動化し、安全な手順で短時間に実行できるようにします。",
    ],
    top=1.6,
    box_height=4.8,
)
add_footer(slide)

# ---------- Slide 5: 期待できる効果 ----------
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_title(slide, "期待できる効果")
cards = [
    (0.8, 1.7, "安全性", "長期の鍵を減らし、漏えいリスクを低減", GREEN),
    (3.8, 1.7, "スピード", "リリースまでの時間を短縮", BLUE),
    (6.8, 1.7, "コスト", "運用工数を減らし、無駄な作業を削減", ORANGE),
    (9.8, 1.7, "監査対応", "ログをまとめて確認しやすくする", NAVY),
]
for x, y, title, desc, color in cards:
    add_label_box(slide, x, y, 2.4, 1.1, title, color, color, 18, True, WHITE)
    add_small_text(
        slide,
        x,
        y + 1.2,
        2.4,
        0.7,
        desc,
        font_size=12,
        color=DARK_GRAY,
        align=PP_ALIGN.CENTER,
    )

add_label_box(
    slide,
    1.0,
    4.7,
    11.3,
    1.0,
    "導入の狙いは「安全に、早く、無駄なく」運用できる状態に変えることです。",
    LIGHT_BLUE,
    BLUE,
    18,
    True,
    NAVY,
)
add_footer(slide)

# ---------- Slide 6: 現状イメージ（いまの課題） ----------
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_title(slide, "現状イメージ（AWSでの開発・運用）")

# 左: 開発/CI（いま）
add_label_box(
    slide,
    0.6,
    2.1,
    3.0,
    1.2,
    "開発・CI/CD\n(GitHub Actions)",
    ORANGE,
    ORANGE,
    18,
    True,
    WHITE,
)
add_small_text(
    slide,
    0.6,
    3.4,
    3.0,
    0.5,
    "長期キー管理・手作業が残る",
    font_size=12,
    color=GRAY,
    align=PP_ALIGN.CENTER,
)

# 中央: リスク要因（点在・公開・手作業）
# 目立つように枠を赤、塗りは薄め（グレー）
add_label_box(
    slide,
    4.2,
    1.6,
    2.4,
    0.9,
    "別々のログイン・権限",
    LIGHT_GRAY,
    RED,
    14,
    True,
    DARK_GRAY,
)
add_label_box(
    slide,
    6.9,
    1.6,
    2.4,
    0.9,
    "各所に点在する\n鍵・パスワード",
    LIGHT_GRAY,
    RED,
    14,
    True,
    DARK_GRAY,
)
add_label_box(
    slide,
    4.2,
    2.7,
    2.4,
    0.9,
    "公開エンドポイント\n（露出多い）",
    LIGHT_GRAY,
    RED,
    14,
    True,
    DARK_GRAY,
)
add_label_box(
    slide,
    6.9,
    2.7,
    2.4,
    0.9,
    "手作業の切替・長期キー",
    LIGHT_GRAY,
    RED,
    14,
    True,
    DARK_GRAY,
)
add_small_text(
    slide,
    4.2,
    3.7,
    5.1,
    0.4,
    "複雑さ＝ミスと漏えいの入口が増える",
    font_size=12,
    color=GRAY,
    align=PP_ALIGN.CENTER,
)

# 右: 業務アプリ（分散）
add_label_box(
    slide,
    9.0,
    2.1,
    3.4,
    1.2,
    "業務アプリ・データ\n（複数環境に分散）",
    ORANGE,
    ORANGE,
    18,
    True,
    WHITE,
)
add_small_text(
    slide,
    9.0,
    3.4,
    3.4,
    0.5,
    "管理と監査が難しい",
    font_size=12,
    color=GRAY,
    align=PP_ALIGN.CENTER,
)

# 矢印（複雑さを表すために複数）
# 左→中央
add_arrow(slide, 3.3, 2.15, 0.6, 0.25, fill=RED)
add_arrow(slide, 3.3, 1.95, 0.6, 0.25, fill=RED)
add_arrow(slide, 3.3, 2.95, 0.6, 0.25, fill=RED)
# 中央→右（いくつかの箱から）
add_arrow(slide, 8.0, 1.85, 0.8, 0.25, fill=RED)
add_arrow(slide, 8.0, 2.95, 0.8, 0.25, fill=RED)

# 注意メッセージ
add_label_box(
    slide,
    1.0,
    5.6,
    11.3,
    0.9,
    "複雑な構成と点在する機密情報は、運用負荷とセキュリティ事故の可能性を高めます。",
    RGBColor(255, 242, 204),
    ORANGE,
    18,
    True,
    DARK_GRAY,
)
add_footer(slide)

# ---------- Slide 6: 全体フロー図 ----------
slide = prs.slides.add_slide(prs.slide_layouts[6])
# add_title(slide, "全体イメージ（分かりやすい流れ）")
add_title(slide, "将来イメージ（Azure移行後）")

add_label_box(
    slide,
    0.7,
    2.2,
    3.1,
    1.3,
    "開発・CI/CD\n(GitHub Actions)",
    BLUE,
    BLUE,
    18,
    True,
    WHITE,
)
add_small_text(
    slide,
    0.7,
    3.6,
    3.1,
    0.5,
    "短期の認証で安全に接続",
    font_size=12,
    color=GRAY,
    align=PP_ALIGN.CENTER,
)

add_label_box(slide, 4.9, 1.8, 3.5, 1.0, "Azure 共通基盤", NAVY, NAVY, 20, True, WHITE)
add_label_box(
    slide, 4.9, 3.0, 1.65, 1.2, "社員の\nログイン", LIGHT_BLUE, BLUE, 16, True, NAVY
)
add_label_box(
    slide, 6.75, 3.0, 1.65, 1.2, "鍵の保管\n自動管理", LIGHT_BLUE, BLUE, 16, True, NAVY
)
add_small_text(
    slide,
    4.9,
    4.35,
    3.5,
    0.4,
    "ここで認証と機密管理をまとめる",
    font_size=12,
    color=GRAY,
    align=PP_ALIGN.CENTER,
)

add_label_box(
    slide, 9.2, 2.2, 3.1, 1.3, "業務アプリ・データ", GREEN, GREEN, 18, True, WHITE
)
add_small_text(
    slide,
    9.2,
    3.6,
    3.1,
    0.5,
    "安全に接続・データ交換",
    font_size=12,
    color=GRAY,
    align=PP_ALIGN.CENTER,
)

add_arrow(slide, 3.95, 2.55, 0.75, 0.4, fill=BLUE)
add_arrow(slide, 8.55, 2.55, 0.75, 0.4, fill=GREEN)

add_label_box(
    slide,
    1.0,
    5.4,
    11.3,
    0.9,
    "外から直接見えにくい形にして、不要な公開を減らします。",
    RGBColor(255, 242, 204),
    ORANGE,
    18,
    True,
    DARK_GRAY,
)
add_footer(slide)

# ---------- Slide 7: 導入ステップ ----------
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_title(slide, "導入ステップ")

steps = [
    (0.8, "1〜2か月", "現状調査\n試験環境の準備", BLUE),
    (4.3, "3〜6か月", "主要システムの切替\n運用ルール整備", NAVY),
    (7.8, "6〜12か月", "残りの移行\n監視と監査の完成", GREEN),
]
for x, t, d, c in steps:
    add_label_box(slide, x, 2.0, 2.8, 1.0, t, c, c, 18, True, WHITE)
    add_label_box(slide, x, 3.2, 2.8, 1.5, d, LIGHT_GRAY, GRAY, 16, True, DARK_GRAY)

add_arrow(slide, 3.65, 2.35, 0.5, 0.35, fill=GRAY)
add_arrow(slide, 7.15, 2.35, 0.5, 0.35, fill=GRAY)

add_label_box(
    slide,
    0.9,
    5.3,
    11.3,
    0.9,
    "段階的に進めることで、業務への影響を抑えながら移行できます。",
    LIGHT_BLUE,
    BLUE,
    18,
    True,
    NAVY,
)
add_footer(slide)

# ---------- Slide 8: リスクと対策 ----------
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_title(slide, "リスクと対策")

left_x = 0.8
right_x = 6.9
y = 1.7

risks = [
    ("移行中の一時的な負荷", "段階的に移行し、並行運用で停止を避ける。"),
    ("既存システムの改修コスト", "効果の大きいところから着手し、早く成果を出す。"),
    ("運用ルールの変更", "手順書と教育を整え、最初は手厚く支援する。"),
]

for i, (risk, counter) in enumerate(risks):
    yy = y + i * 1.35
    add_label_box(
        slide,
        left_x,
        yy,
        5.1,
        0.5,
        "リスク: " + risk,
        RGBColor(255, 242, 204),
        ORANGE,
        14,
        True,
        DARK_GRAY,
    )
    add_label_box(
        slide,
        right_x,
        yy,
        5.6,
        0.5,
        "対策: " + counter,
        LIGHT_BLUE,
        BLUE,
        14,
        True,
        NAVY,
    )

add_label_box(
    slide,
    0.9,
    6.0,
    11.3,
    0.7,
    "リスクをゼロにするのではなく、影響が小さい進め方にすることが重要です。",
    LIGHT_BLUE,
    BLUE,
    16,
    True,
    NAVY,
)
add_footer(slide)

# ---------- Slide 9: 決裁依頼 ----------
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_title(slide, "ご決裁いただきたいこと")

add_bullets(
    slide,
    "",
    [
        "移行プロジェクトの開始承認",
        "初期調査と試験導入の予算承認",
        "関係部門の担当者アサイン",
    ],
    top=1.7,
    box_height=2.0,
)

add_label_box(
    slide,
    1.0,
    4.2,
    11.2,
    1.2,
    "初期投資で、将来の運用負荷とセキュリティリスクを下げる判断をお願いします。",
    NAVY,
    NAVY,
    18,
    True,
    WHITE,
)
add_label_box(
    slide,
    1.0,
    5.8,
    11.2,
    0.8,
    "次のステップ: 詳細見積・移行計画・試験導入",
    RGBColor(255, 242, 204),
    ORANGE,
    18,
    True,
    DARK_GRAY,
)
add_footer(slide)

prs.save(OUTPUT_FILE)
print(f"PowerPoint file created: {OUTPUT_FILE}")
