"""
generate_slides_v10.py
──────────────────────────────────────────────────────────────────────────
多节点信号级高效相参合成机理与方法  ·  研究方案框架图（单页 PPT）

布局（16:9  13.33" × 7.50"）：
  ┌────────────────────────────────────────────────────────────────────┐
  │  标题栏（与 v8 同款）                                              │
  ├────────────────────────────────────────────────────────────────────┤
  │  研究动机与背景：[应用需求] ──► [工程挑战] ──► [科学空白]          │
  ├──────────┬──────────────────────────────────────────────┬──────────┤
  │          │ 科学问题A  │ 科学问题B  │ 科学问题C           │          │
  │   研究   │   增益律   │  相位误差  │  融合准则           │   预期   │
  │   目标   │    ▼       │    ▼       │    ▼                │   成果   │
  │          │ 内容一     │ 内容二     │ 内容三              │          │
  │          │ 量化模型   │ 动态估计   │ 自适应合并          │          │
  │          │    ▼       │    ▼       │    ▼                │          │
  │          │ 技术一     │ 技术二     │ 技术三              │          │
  ├──────────┴──────────────────────────────────────────────┴──────────┤
  │  [底部横条] 高效性指标体系  G≥14dB·σ_φ<π/8·延迟<9ms              │
  ├────────────────────────────────────────────────────────────────────┤
  │  页脚                                                              │
  └────────────────────────────────────────────────────────────────────┘

风格：与 generate_slides_v8.py 完全一致。

运行：
    python generate_slides_v10.py
输出：
    coherent_synthesis_slides_v10.pptx
"""

import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# ━━━━━━━━━━━━━━━━━━━━━━ 调色板（与 v8 完全一致）━━━━━━━━━━━━━━━━━━━━━━
BG          = RGBColor(0xF4, 0xF7, 0xFC)
CARD_BG     = RGBColor(0xFF, 0xFF, 0xFF)
TITLE_BAR   = RGBColor(0x1F, 0x38, 0x64)
TITLE_BAR2  = RGBColor(0x2E, 0x74, 0xB5)
DARK_TEXT   = RGBColor(0x1A, 0x1A, 0x3A)
WHITE       = RGBColor(0xFF, 0xFF, 0xFF)
CARD_BORDER = RGBColor(0xC8, 0xD8, 0xF0)
SEPARATOR   = RGBColor(0xC8, 0xD8, 0xF0)
FOOTER_BG   = RGBColor(0xE8, 0xF0, 0xFB)
COL_NAVY    = RGBColor(0x0D, 0x2A, 0x6E)

# 研究背景区（橙/黄色系）
BG1 = RGBColor(0xC0, 0x50, 0x00)   # 应用需求
BG2 = RGBColor(0x8B, 0x45, 0x00)   # 工程挑战
BG3 = RGBColor(0x6E, 0x25, 0x00)   # 科学空白

# 研究目标/成果侧栏（深蓝紫）
SIDE_COL    = RGBColor(0x0D, 0x2A, 0x6E)
SIDE_BG     = RGBColor(0xE4, 0xEA, 0xF8)

# 科学问题（橙棕）
SCI_COLS = [
    RGBColor(0xB5, 0x4A, 0x00),
    RGBColor(0x8B, 0x2E, 0x00),
    RGBColor(0x6E, 0x20, 0x00),
]

# 研究内容（绿青）
RC_COLS = [
    RGBColor(0x15, 0x7A, 0x35),
    RGBColor(0x00, 0x7C, 0x78),
    RGBColor(0x0E, 0x6B, 0x5E),
]
RC_BG   = [
    RGBColor(0xEA, 0xF6, 0xEF),
    RGBColor(0xE6, 0xF5, 0xF4),
    RGBColor(0xE8, 0xF4, 0xF2),
]

# 关键技术（蓝紫）
KT_COLS = [
    RGBColor(0x1F, 0x6B, 0xC4),
    RGBColor(0x6A, 0x2B, 0xB5),
    RGBColor(0x1F, 0x38, 0x64),
]
KT_BG   = [
    RGBColor(0xEB, 0xF0, 0xFF),
    RGBColor(0xF2, 0xEB, 0xFF),
    RGBColor(0xE8, 0xEC, 0xF8),
]

# 底部指标栏
BOT_COL = RGBColor(0x0D, 0x2A, 0x6E)

# ━━━━━━━━━━━━━━━━━━━━━━ 尺寸常量 ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
SW      = Inches(13.33)
SH      = Inches(7.50)
MARGIN  = Inches(0.32)
TITLE_H = Inches(0.82)
CT      = TITLE_H + Inches(0.14)
FTR_H   = Inches(0.48)
FOOTER_Y = SH - FTR_H
CW      = SW - 2 * MARGIN


# ━━━━━━━━━━━━━━━━━━━━━━ 工具函数（与 v8 一致）━━━━━━━━━━━━━━━━━━━━━━━

def set_bg(slide, color):
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = color


def rect(slide, l, t, w, h, fill, lc=None, lw=Pt(0.75), rnd=False):
    s = slide.shapes.add_shape(5 if rnd else 1, l, t, w, h)
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    if lc:
        s.line.color.rgb = lc
        s.line.width = lw
    else:
        s.line.fill.background()
    return s


def hline(slide, x1, y, x2, color=SEPARATOR, width=Pt(0.6)):
    c = slide.shapes.add_connector(1, x1, y, x2, y)
    c.line.color.rgb = color
    c.line.width = width


def vline(slide, x, y1, y2, color=SEPARATOR, width=Pt(0.6)):
    c = slide.shapes.add_connector(1, x, y1, x, y2)
    c.line.color.rgb = color
    c.line.width = width


def tb(slide, text, l, t, w, h,
       sz=10, bold=False, color=DARK_TEXT,
       align=PP_ALIGN.LEFT, italic=False, wrap=True):
    box = slide.shapes.add_textbox(l, t, w, h)
    tf  = box.text_frame
    tf.word_wrap = wrap
    p   = tf.paragraphs[0]
    p.alignment = align
    r   = p.add_run()
    r.text = text
    r.font.size   = Pt(sz)
    r.font.bold   = bold
    r.font.italic = italic
    r.font.color.rgb = color
    return box


def tb_mp(slide, lines, l, t, w, h,
          sz=9, bold=False, color=DARK_TEXT,
          align=PP_ALIGN.LEFT, wrap=True):
    """Multi-paragraph text box."""
    box = slide.shapes.add_textbox(l, t, w, h)
    tf  = box.text_frame
    tf.word_wrap = wrap
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        r = p.add_run()
        r.text = ln
        r.font.size  = Pt(sz)
        r.font.bold  = bold
        r.font.color.rgb = color
    return box


def circle_badge(slide, cx, cy, r_inch, fill, text,
                 sz=10, bold=True, tc=WHITE):
    d = Inches(r_inch * 2)
    s = slide.shapes.add_shape(9,
                                cx - Inches(r_inch),
                                cy - Inches(r_inch), d, d)
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    s.line.fill.background()
    tf = s.text_frame
    tf.word_wrap = False
    p  = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r  = p.add_run()
    r.text = text
    r.font.size  = Pt(sz)
    r.font.bold  = bold
    r.font.color.rgb = tc


def title_bar(slide, title_main, title_sub, page, total=1):
    rect(slide, Inches(0), Inches(0), Inches(0.18), TITLE_H, COL_NAVY)
    rect(slide, Inches(0.18), Inches(0), SW - Inches(0.18), TITLE_H, TITLE_BAR)
    hline(slide, Inches(0), TITLE_H, SW, TITLE_BAR2, Pt(2.5))
    tb(slide, title_main,
       Inches(0.36), Inches(0.10), Inches(10.5), Inches(0.38),
       sz=20, bold=True, color=WHITE)
    tb(slide, title_sub,
       Inches(0.36), Inches(0.50), Inches(10.5), Inches(0.26),
       sz=10.5, color=RGBColor(0xB0, 0xCC, 0xF0))
    circle_badge(slide, SW - Inches(0.52), TITLE_H / 2, 0.28,
                 TITLE_BAR2, f"{page}/{total}", sz=11)


def footer(slide, text):
    rect(slide, Inches(0), FOOTER_Y, SW, FTR_H, FOOTER_BG)
    hline(slide, MARGIN, FOOTER_Y, SW - MARGIN, TITLE_BAR2, Pt(1.5))
    rect(slide, MARGIN, FOOTER_Y + Inches(0.10),
         Inches(0.05), Inches(0.28), TITLE_BAR2)
    tb(slide, "  " + text,
       MARGIN + Inches(0.10), FOOTER_Y + Inches(0.09),
       CW - Inches(0.14), Inches(0.34),
       sz=8.8, italic=True, color=COL_NAVY)


# ━━━━━━━━━━━━━━━━━━ 带标题色块的区块工具 ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def section_card(slide, l, t, w, h, hdr_color, bg_color,
                 badge, title_line1, title_line2,
                 body_lines, hdr_h=Inches(0.36),
                 title_sz=10, body_sz=8.8, border_color=None):
    """
    通用区块卡片：彩色头带 + 浅色正文区 + 左侧强调竖条
    badge      : 圆形编号标（如 "A"）
    title_line1: 粗标题
    title_line2: 副标题（可空）
    body_lines : [(text, is_accent)] list
    """
    ACC_W = Inches(0.09)
    # 阴影
    rect(slide, l + Inches(0.025), t + Inches(0.025), w, h,
         RGBColor(0xC0, 0xC8, 0xD8))
    # 主背景
    bc = border_color or hdr_color
    rect(slide, l, t, w, h, bg_color, lc=bc, lw=Pt(0.8))
    # 头带
    rect(slide, l, t, w, hdr_h, hdr_color)
    # 左侧强调竖条（正文区）
    rect(slide, l, t + hdr_h, ACC_W, h - hdr_h, hdr_color)
    # 编号圆
    if badge:
        circle_badge(slide, l + Inches(0.22), t + hdr_h / 2,
                     0.155, WHITE, badge, sz=8.5, tc=hdr_color)
    bx = l + (Inches(0.45) if badge else Inches(0.12))
    bw = w - bx + l - Inches(0.08)
    tb(slide, title_line1, bx, t + Inches(0.04),
       bw, Inches(0.22), sz=title_sz, bold=True, color=WHITE)
    if title_line2:
        tb(slide, title_line2, bx, t + hdr_h - Inches(0.14),
           bw, Inches(0.14), sz=8, color=RGBColor(0xCC, 0xEA, 0xFF))
    # 正文
    pad_l = l + Inches(0.16)
    pad_w = w - Inches(0.22)
    yy    = t + hdr_h + Inches(0.055)
    lh    = body_sz * 1.75 * 914.4 / 72   # EMU per line
    for (txt, accent) in body_lines:
        clr = hdr_color if accent else DARK_TEXT
        tb(slide, txt, pad_l, yy, pad_w, lh,
           sz=body_sz, color=clr, bold=accent)
        yy += lh + Inches(0.012)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━ MAIN ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

prs = Presentation()
prs.slide_width  = SW
prs.slide_height = SH
blank = prs.slide_layouts[6]

s = prs.slides.add_slide(blank)
set_bg(s, BG)

# ── 1. 标题栏 ──────────────────────────────────────────────────────
title_bar(s,
          "多节点信号级高效相参合成机理与方法  ·  研究方案框架图",
          "Research Framework: Mechanism & Methods for Efficient Multi-Node Signal-Level Coherent Synthesis",
          1)

# ── 2. 计算全局布局 ────────────────────────────────────────────────
BOT_STRIP_H = Inches(0.38)
BOT_Y       = FOOTER_Y - BOT_STRIP_H - Inches(0.04)

TOP_Y    = CT                          # 内容顶部
AVAIL_H  = BOT_Y - TOP_Y - Inches(0.04)

# 各行高比例：背景带 0.60  科学问题 0.78  研究内容 1.72  关键技术 1.05
ROW_BG_H = Inches(0.60)
ROW_SCI_H = Inches(0.78)
ROW_RC_H  = Inches(1.72)
ROW_KT_H  = Inches(1.05)
ROW_GAP   = Inches(0.095)

y_bg  = TOP_Y
y_sci = y_bg  + ROW_BG_H  + ROW_GAP
y_rc  = y_sci + ROW_SCI_H + ROW_GAP
y_kt  = y_rc  + ROW_RC_H  + ROW_GAP

# 侧栏宽度（左"研究目标"，右"预期成果"）
SIDE_W = Inches(1.12)
SIDE_GAP = Inches(0.10)
INNER_L  = MARGIN + SIDE_W + SIDE_GAP
INNER_W  = CW - 2 * (SIDE_W + SIDE_GAP)
COL_GAP  = Inches(0.10)
COL_W    = (INNER_W - 2 * COL_GAP) / 3
col_xs   = [INNER_L + i * (COL_W + COL_GAP) for i in range(3)]
RIGHT_X  = MARGIN + CW - SIDE_W

# ── 3. 研究背景带（顶部横条）──────────────────────────────────────
BG_STEPS = [
    ("应用需求牵引", BG1, [
        "远程组网精确打击",
        "低截获·高机动探测",
        "多平台协同感知",
    ]),
    ("工程技术挑战", BG2, [
        "严苛时频同步约束",
        "高动态时变相位误差",
        "多平台异步融合难题",
    ]),
    ("核心科学空白", BG3, [
        "增益退化律未量化",
        "相位估计CRLB边界不清",
        "最优融合准则缺失",
    ]),
]
n_bg    = len(BG_STEPS)
arr_w   = Inches(0.13)
bg_usable = CW - arr_w * (n_bg - 1)
bg_box_w  = bg_usable / n_bg
bx = MARGIN
for i, (lbl, col, items) in enumerate(BG_STEPS):
    rect(s, bx, y_bg, bg_box_w, ROW_BG_H, col, rnd=True)
    # 标题
    tb(s, lbl, bx + Inches(0.10), y_bg + Inches(0.05),
       bg_box_w - Inches(0.14), Inches(0.20),
       sz=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    # 条目（竖向均分）
    item_y = y_bg + Inches(0.28)
    item_h = Inches(0.16)
    for it in items:
        tb(s, "• " + it, bx + Inches(0.12), item_y,
           bg_box_w - Inches(0.18), item_h,
           sz=8.2, color=WHITE)
        item_y += item_h
    if i < n_bg - 1:
        tb(s, "►", bx + bg_box_w, y_bg + (ROW_BG_H - Inches(0.20)) / 2,
           arr_w, Inches(0.20), sz=10, color=COL_NAVY,
           align=PP_ALIGN.CENTER)
    bx += bg_box_w + arr_w

# 背景带下方箭头指向"科学问题"
for ci in range(3):
    mid_x = col_xs[ci] + COL_W / 2
    vline(s, mid_x, y_bg + ROW_BG_H, y_sci - Inches(0.01),
          TITLE_BAR2, Pt(1.2))
    tb(s, "▼", mid_x - Inches(0.08), y_sci - Inches(0.135),
       Inches(0.16), Inches(0.14), sz=8.5,
       color=TITLE_BAR2, align=PP_ALIGN.CENTER)

# ── 4. 左侧"研究目标"侧栏 ─────────────────────────────────────────
SIDE_H = y_kt + ROW_KT_H - y_sci
rect(s, MARGIN + Inches(0.025), y_sci + Inches(0.025),
     SIDE_W, SIDE_H, RGBColor(0xB0, 0xB8, 0xD8))   # 阴影
rect(s, MARGIN, y_sci, SIDE_W, SIDE_H, SIDE_COL, rnd=True)
rect(s, MARGIN, y_sci, Inches(0.10), SIDE_H,
     RGBColor(0x50, 0x90, 0xFF))   # 左高亮竖条
tb(s, "研\n究\n目\n标",
   MARGIN + Inches(0.14), y_sci + (SIDE_H - Inches(0.80)) / 2,
   SIDE_W - Inches(0.16), Inches(0.80),
   sz=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
tb(s, "Research\nObjectives",
   MARGIN + Inches(0.06), y_sci + SIDE_H - Inches(0.42),
   SIDE_W - Inches(0.08), Inches(0.40),
   sz=7, color=RGBColor(0xA0, 0xC4, 0xFF),
   align=PP_ALIGN.CENTER)

# ── 5. 右侧"预期成果"侧栏 ─────────────────────────────────────────
RESULT_ITEMS = [
    "G≥14 dB（N=6）",
    "σ_est < π/8（SNR≥6dB）",
    "延迟 < 9 ms（实时）",
    "发表 SCI 论文 2~3 篇",
    "申请发明专利 1~2 项",
]
rect(s, RIGHT_X + Inches(0.025), y_sci + Inches(0.025),
     SIDE_W, SIDE_H, RGBColor(0xB0, 0xB8, 0xD8))
rect(s, RIGHT_X, y_sci, SIDE_W, SIDE_H, SIDE_COL, rnd=True)
rect(s, RIGHT_X + SIDE_W - Inches(0.10), y_sci,
     Inches(0.10), SIDE_H, RGBColor(0x50, 0x90, 0xFF))
tb(s, "预\n期\n成\n果",
   RIGHT_X + Inches(0.02), y_sci + Inches(0.04),
   SIDE_W - Inches(0.14), Inches(0.80),
   sz=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
item_y2 = y_sci + Inches(0.90)
for ri in RESULT_ITEMS:
    tb(s, "✓ " + ri, RIGHT_X + Inches(0.06), item_y2,
       SIDE_W - Inches(0.12), Inches(0.155),
       sz=7.8, color=WHITE)
    item_y2 += Inches(0.155)

# ── 6. 科学问题行（3格）──────────────────────────────────────────
SCI_DATA = [
    ("A", "科学问题 A",
     "增益退化规律（SNR增益律）",
     [("增益G与同步误差(σ_τ,σ_φ,σ_f)的定量关系？", True),
      ("• N节点理论上限 G=N² 在工程约束下如何退化？", False),
      ("• 效能保持90%的最大容忍误差边界是多少？", False)]),
    ("B", "科学问题 B",
     "时变相位误差建模与可估计性",
     [("高动态场景下相位漂移率的最优估计精度边界？", True),
      ("• Δφ_k(t) 的非平稳模型如何准确刻画？", False),
      ("• 无导频条件下 CRLB 与帧长/SNR 的约束关系？", False)]),
    ("C", "科学问题 C",
     "信号质量自适应最优融合准则",
     [("SNR、相位一致性、帧长联合约束下最优融合？", True),
      ("• 异构节点质量差异大时如何稳健合并？", False),
      ("• 层次化处理框架的实时复杂度如何优化？", False)]),
]

for ci, (badge, sci_title, sci_sub, items) in enumerate(SCI_DATA):
    section_card(
        s, col_xs[ci], y_sci, COL_W, ROW_SCI_H,
        SCI_COLS[ci], RGBColor(0xFF, 0xF4, 0xEE),
        badge, sci_title, sci_sub, items,
        hdr_h=Inches(0.34), title_sz=9.8, body_sz=8.2,
        border_color=SCI_COLS[ci]
    )

# 科学问题→研究内容 箭头
for ci in range(3):
    mid_x = col_xs[ci] + COL_W / 2
    vline(s, mid_x, y_sci + ROW_SCI_H, y_rc - Inches(0.01),
          TITLE_BAR2, Pt(1.2))
    tb(s, "▼", mid_x - Inches(0.08), y_rc - Inches(0.135),
       Inches(0.16), Inches(0.14), sz=8.5,
       color=TITLE_BAR2, align=PP_ALIGN.CENTER)

# ── 7. 研究内容行（3卡片）────────────────────────────────────────
RC_DATA = [
    ("一", "研究内容一",
     "增益退化端到端量化模型",
     [("• 建立时延/相位/频偏→合并增益G的统一传播链路", False),
      ("G_eff = G·exp(−σ²_φ)·sinc²(σ_τB)·sinc²(σ_fNcPRI)", True),
      ('• 量化"效能90%保持"误差边界：工程容差准则', False),
      ("• N=2,4,6节点仿真验证退化曲线与灵敏度分析", False),
      ("• 推导 BCRLB 效能下界·对比蒙特卡洛仿真", False),
      ("• 输出 G_eff(σ_τ,σ_φ,σ_f) 三维效能图谱", False)]),
    ("二", "研究内容二",
     "高动态时变相位误差建模与实时估计",
     [("• 建立非平稳相位模型：运动/振荡/多普勒三分量叠加", False),
      ("Δφ_k(t) = φ_delay(t) + φ_osc(t) + φ_Doppler(t)", True),
      ("• 无导频自校准：互相关亚像素峰值→EKF实时跟踪", False),
      ("• 推导最小可辨识帧长 N_c,min 与 SNR 联合约束", False),
      ("• 自适应帧长选择准则：精度-延迟动态权衡", False),
      ("• SNR≥6dB时 σ_est < π/8，满足 G_eff ≥ 90%G", False)]),
    ("三", "研究内容三",
     "质量自适应MRC合并与层次化处理框架",
     [("• MRC权值：w_k = SNR_k·coh_k·ph_cons_k 三因子联合估计", False),
      ("w_k ∝ SNR_k · coh_k · ph_cons_k  (在线迭代更新)", True),
      ("• 稳健机制：相位估计失败→自动降权，防止输出突变", False),
      ("• 分层策略：强目标SIC对消→弱目标MTI→统一MRC合并", False),
      ("• 子带分解+稀疏采样：复杂度降一量级→延迟<9 ms", False),
      ("• 数据集：MATLAB端到端仿真+半实物验证平台", False)]),
]

for ci, (badge, rc_title, rc_sub, items) in enumerate(RC_DATA):
    section_card(
        s, col_xs[ci], y_rc, COL_W, ROW_RC_H,
        RC_COLS[ci], RC_BG[ci],
        badge, rc_title, rc_sub, items,
        hdr_h=Inches(0.36), title_sz=10, body_sz=8.5,
        border_color=RC_COLS[ci]
    )

# 研究内容→关键技术 箭头
for ci in range(3):
    mid_x = col_xs[ci] + COL_W / 2
    vline(s, mid_x, y_rc + ROW_RC_H, y_kt - Inches(0.01),
          TITLE_BAR2, Pt(1.2))
    tb(s, "▼", mid_x - Inches(0.08), y_kt - Inches(0.135),
       Inches(0.16), Inches(0.14), sz=8.5,
       color=TITLE_BAR2, align=PP_ALIGN.CENTER)

# ── 8. 关键技术行（3格）─────────────────────────────────────────
KT_DATA = [
    ("①", "关键技术一",
     "误差传播链路建模与效能边界推导",
     [("• 联合误差传播理论 + 贝叶斯效能下界（BCRLB）", False),
      ("• G_eff vs (σ_τ,σ_φ,σ_f) 三维效能图谱及工程容差曲线", False),
      ("• 预期：N=6时 G_eff≥14dB（误差容限内，置信度≥90%）", True)]),
    ("②", "关键技术二",
     "EKF 非平稳相位在线估计与自校准",
     [("• 状态量[Δφ_k, φ̇_k]，过程噪声由平台运动学参数驱动", False),
      ("• 互相关亚像素峰值提供Δτ观测，EKF步长=1 CPI", False),
      ("• 预期：SNR≥6dB时 σ_est<π/8，满足G_eff≥90%G", True)]),
    ("③", "关键技术三",
     "分层相参策略与实时计算优化",
     [("• 两轮融合：CFAR强目标SIC对消→三MTI并行+MRC合并", False),
      ("• 子带分解压缩至<5ms；精度损失<0.3dB（仿真验证）", False),
      ("• 预期：Nc=64,Nsc=1024,N=6链路→总延迟<9ms", True)]),
]

for ci, (badge, kt_title, kt_sub, items) in enumerate(KT_DATA):
    section_card(
        s, col_xs[ci], y_kt, COL_W, ROW_KT_H,
        KT_COLS[ci], KT_BG[ci],
        badge, kt_title, kt_sub, items,
        hdr_h=Inches(0.34), title_sz=9.8, body_sz=8.5,
        border_color=KT_COLS[ci]
    )

# ── 9. 底部指标横栏 ──────────────────────────────────────────────
rect(s, Inches(0), BOT_Y, SW, BOT_STRIP_H, BOT_COL)
# 左侧图标区
rect(s, Inches(0), BOT_Y, Inches(0.60), BOT_STRIP_H,
     RGBColor(0x2E, 0x74, 0xB5))
tb(s, "🎯",
   Inches(0.06), BOT_Y + Inches(0.05),
   Inches(0.48), BOT_STRIP_H - Inches(0.08),
   sz=16, align=PP_ALIGN.CENTER, color=WHITE)
# 标签
tb(s, "高效性量化指标体系  Quantitative Efficiency Metrics",
   Inches(0.70), BOT_Y + Inches(0.04),
   Inches(3.00), Inches(0.20),
   sz=9.5, bold=True, color=WHITE)
# 各指标
metrics = [
    ("G ≥ 14 dB",        "N=6时效能指标"),
    ("σ_est < π/8",      "相位误差上限"),
    ("延迟 < 9 ms",       "实时处理指标"),
    ("SCI ×2~3篇",        "学术成果指标"),
    ("发明专利 ×1~2项",   "工程转化成果"),
]
mx = Inches(0.70)
m_w = (SW - mx - Inches(0.30)) / len(metrics)
for val, lbl in metrics:
    tb(s, val, mx, BOT_Y + Inches(0.24),
       m_w, Inches(0.18), sz=9.5, bold=True,
       color=RGBColor(0xFF, 0xE0, 0x80), align=PP_ALIGN.CENTER)
    tb(s, lbl, mx, BOT_Y + Inches(0.22 + 0.16),
       m_w, Inches(0.14), sz=7.5,
       color=RGBColor(0xC0, 0xD8, 0xFF), align=PP_ALIGN.CENTER)
    mx += m_w

# 分隔竖线
mx2 = Inches(0.70) + m_w
for _ in range(len(metrics) - 1):
    vline(s, mx2 - Inches(0.02), BOT_Y + Inches(0.08),
          BOT_Y + BOT_STRIP_H - Inches(0.06),
          RGBColor(0x60, 0x80, 0xC0), Pt(0.5))
    mx2 += m_w

# ── 10. 页脚 ─────────────────────────────────────────────────────
footer(s,
       "研究路径：工程需求牵引 → 三大科学问题凝练 → 三项研究内容攻关 → 三项关键技术突破 → "
       "理论 + 算法 + 工程验证全链路闭环，最终实现 N=6 节点 G_eff≥14dB·延迟<9ms 目标")

# ── 11. 保存 ─────────────────────────────────────────────────────
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                   "coherent_synthesis_slides_v10.pptx")
prs.save(OUT)
print(f"✅  已生成：{OUT}")
