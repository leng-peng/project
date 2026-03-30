"""
generate_slides_v6.py
─────────────────────────────────────────────────────────────────────────
空海平台多节点协同探测 — 相参合成两页 PPT（v6）

  Page 1（机理分析 · 重绘版）：
    突出「距离多普勒域信号级融合」与「数据转换」两大核心机理
    布局：
      ① 顶部全宽信号处理流程图（原始回波→脉压→RD图→数据转换→融合→检测）
      ② 主内容区 · 左列「数据转换」：3个子卡片
         - 双基地时延配准（等效斜距折叠）
         - 多普勒域统一（速度转换因子）
         - 坐标系配准（ECEF/ENU 运动补偿）
      ③ 主内容区 · 右列「距离多普勒域信号级融合」：3个子卡片
         - RD图生成与信号模型
         - 相参合并公式（MRC权值）
         - 增益退化律（效能分析）

  Page 2（方法设计，与 v5 保持一致）

运行：
    python generate_slides_v6.py
输出：
    coherent_synthesis_slides_v6.pptx
"""

import os

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# ━━━━━━━━━━━━━━━━━━━━━━ 调色板 ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
BG           = RGBColor(0xF4, 0xF7, 0xFC)
CARD_BG      = RGBColor(0xFF, 0xFF, 0xFF)
TITLE_BAR    = RGBColor(0x1F, 0x38, 0x64)
TITLE_BAR2   = RGBColor(0x2E, 0x74, 0xB5)
DARK_TEXT    = RGBColor(0x1A, 0x1A, 0x3A)
MID_TEXT     = RGBColor(0x3D, 0x3D, 0x6B)
WHITE        = RGBColor(0xFF, 0xFF, 0xFF)
CARD_BORDER  = RGBColor(0xC8, 0xD8, 0xF0)
SEPARATOR    = RGBColor(0xC8, 0xD8, 0xF0)
FOOTER_BG    = RGBColor(0xE8, 0xF0, 0xFB)

# 主题色 — 数据转换（橙色系）
DT_HDR    = RGBColor(0xB5, 0x4A, 0x00)   # 深橙（大区标题）
DT_SUB1   = RGBColor(0xC8, 0x62, 0x10)   # 子块1
DT_SUB2   = RGBColor(0xD4, 0x78, 0x20)   # 子块2
DT_SUB3   = RGBColor(0xE0, 0x90, 0x30)   # 子块3

# 主题色 — 信号级融合（蓝色系）
RD_HDR    = RGBColor(0x0D, 0x2A, 0x6E)   # 深蓝（大区标题）
RD_SUB1   = RGBColor(0x1F, 0x6B, 0xC4)   # 子块1
RD_SUB2   = RGBColor(0x00, 0x7C, 0x78)   # 子块2（青）
RD_SUB3   = RGBColor(0x6A, 0x2B, 0xB5)   # 子块3（紫）

# 流程图色
PIPE_COLS = [
    RGBColor(0x1F, 0x6B, 0xC4),
    RGBColor(0x1F, 0x6B, 0xC4),
    RGBColor(0x1F, 0x6B, 0xC4),
    RGBColor(0xB5, 0x4A, 0x00),   # 数据转换 — 橙色
    RGBColor(0x0D, 0x2A, 0x6E),   # 信号级融合 — 深蓝
    RGBColor(0x00, 0x7C, 0x78),
]

COL_NAVY  = RGBColor(0x0D, 0x2A, 0x6E)
FORMULA_BG = RGBColor(0xE8, 0xF2, 0xFF)

# 方法页色（与v5一致）
PIPE_BLUE   = RGBColor(0x1F, 0x6B, 0xC4)
PIPE_RED    = RGBColor(0xA8, 0x12, 0x12)
PIPE_TEAL   = RGBColor(0x00, 0x7C, 0x78)
PIPE_PURPLE = RGBColor(0x6A, 0x2B, 0xB5)
PIPE_GREEN  = RGBColor(0x15, 0x7A, 0x35)
COL_BLUE    = RGBColor(0x1F, 0x6B, 0xC4)
COL_TEAL    = RGBColor(0x00, 0x7C, 0x78)
COL_PURPLE  = RGBColor(0x6A, 0x2B, 0xB5)
COL_ORANGE  = RGBColor(0xC8, 0x55, 0x0A)

# ━━━━━━━━━━━━━━━━━━━━━━ 尺寸常量 ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
SW       = Inches(13.33)
SH       = Inches(7.50)
MARGIN   = Inches(0.32)
TITLE_H  = Inches(0.82)
CT       = TITLE_H + Inches(0.16)
FOOTER_Y = SH - Inches(0.50)
CB       = FOOTER_Y - Inches(0.04)
CW       = SW - 2 * MARGIN
COL_H    = CB - CT


# ━━━━━━━━━━━━━━━━━━━━━━ 工具函数 ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

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


def line(slide, x1, y1, x2, y2=None, color=SEPARATOR, width=Pt(0.6)):
    if y2 is None:
        y2 = y1
    c = slide.shapes.add_connector(1, x1, y1, x2, y2)
    c.line.color.rgb = color
    c.line.width = width


def tb(slide, text, l, t, w, h,
       sz=11, bold=False, color=DARK_TEXT,
       align=PP_ALIGN.LEFT, italic=False, wrap=True):
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.size = Pt(sz)
    r.font.bold = bold
    r.font.italic = italic
    r.font.color.rgb = color
    return box


def tb_lines(slide, lines_list, l, t, w, h, wrap=True):
    """lines_list: [(text, sz, bold, color, align)]"""
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = wrap
    first = True
    for (txt, sz, bold, col, align) in lines_list:
        if first:
            p = tf.paragraphs[0]
            first = False
        else:
            p = tf.add_paragraph()
        p.alignment = align
        r = p.add_run()
        r.text = txt
        r.font.size = Pt(sz)
        r.font.bold = bold
        r.font.color.rgb = col
    return box


def circle_badge(slide, cx, cy, r_inch, fill, text, sz=11, bold=True, tc=WHITE):
    d = Inches(r_inch * 2)
    s = slide.shapes.add_shape(9, cx - Inches(r_inch), cy - Inches(r_inch), d, d)
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    s.line.fill.background()
    tf = s.text_frame
    tf.word_wrap = False
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = text
    r.font.size = Pt(sz)
    r.font.bold = bold
    r.font.color.rgb = tc


def formula_box(slide, l, t, w, h, text, sz=10):
    rect(slide, l, t, w, h, FORMULA_BG, lc=TITLE_BAR2, lw=Pt(0.8))
    tb(slide, text, l + Inches(0.10), t + Inches(0.06),
       w - Inches(0.18), h - Inches(0.10),
       sz=sz, color=COL_NAVY, bold=True)


# ─── 标题栏 ────────────────────────────────────────────────────────
def title_bar(slide, title_main, title_sub, page, total=2):
    rect(slide, Inches(0), Inches(0), Inches(0.18), TITLE_H, COL_NAVY)
    rect(slide, Inches(0.18), Inches(0), SW - Inches(0.18), TITLE_H, TITLE_BAR)
    line(slide, Inches(0), TITLE_H, SW, TITLE_H, TITLE_BAR2, Pt(2.5))
    tb(slide, title_main,
       Inches(0.36), Inches(0.10), Inches(10.0), Inches(0.38),
       sz=20, bold=True, color=WHITE)
    tb(slide, title_sub,
       Inches(0.36), Inches(0.50), Inches(10.0), Inches(0.26),
       sz=10.5, color=RGBColor(0xB0, 0xCC, 0xF0))
    circle_badge(slide, SW - Inches(0.52), TITLE_H / 2, 0.28,
                 TITLE_BAR2, f"{page}/{total}", sz=11)


# ─── 页脚 ──────────────────────────────────────────────────────────
def footer(slide, text):
    rect(slide, Inches(0), FOOTER_Y, SW, SH - FOOTER_Y, FOOTER_BG)
    line(slide, MARGIN, FOOTER_Y, SW - MARGIN, FOOTER_Y, TITLE_BAR2, Pt(1.5))
    rect(slide, MARGIN, FOOTER_Y + Inches(0.10), Inches(0.05), Inches(0.28), TITLE_BAR2)
    tb(slide, "  " + text,
       MARGIN + Inches(0.10), FOOTER_Y + Inches(0.09),
       CW - Inches(0.14), Inches(0.34),
       sz=8.8, italic=True, color=COL_NAVY)


# ─── 小节子卡片 ────────────────────────────────────────────────────
def sub_card(slide, l, t, w, h, accent, title, items,
             badge=None, title_sz=10.8, body_sz=9.2):
    """
    Sub-card: colored header band + white body.
    items: [(text, is_formula)]
    """
    HDR_H = Inches(0.35)
    # Shadow
    rect(slide, l + Inches(0.03), t + Inches(0.03), w, h,
         RGBColor(0xC8, 0xD4, 0xEC))
    # Body
    rect(slide, l, t, w, h, CARD_BG, lc=CARD_BORDER, lw=Pt(0.7))
    # Header
    rect(slide, l, t, w, HDR_H, accent)
    if badge:
        circle_badge(slide, l + Inches(0.20), t + HDR_H / 2, 0.15,
                     WHITE, badge, sz=8.5, tc=accent)
    hx = l + (Inches(0.44) if badge else Inches(0.12))
    tb(slide, title, hx, t + Inches(0.05),
       w - hx + l - Inches(0.08), HDR_H - Inches(0.08),
       sz=title_sz, bold=True, color=WHITE)
    # Body content
    yy = t + HDR_H + Inches(0.06)
    item_h = body_sz * 1.80 * 914.4 / 72
    for (txt, is_formula) in items:
        clr = COL_NAVY if is_formula else DARK_TEXT
        prefix = "   " if is_formula else "• "
        tb(slide, prefix + txt,
           l + Inches(0.10), yy,
           w - Inches(0.16), item_h,
           sz=body_sz, color=clr)
        yy += item_h + Inches(0.01)


# ─── 大区标题横条 ──────────────────────────────────────────────────
def section_header(slide, l, t, w, h, fill, icon, title,
                   sub=None, title_sz=13, sub_sz=9):
    rect(slide, l, t, w, h, fill)
    # Left icon block
    rect(slide, l, t, Inches(0.08), h, RGBColor(0xFF, 0xFF, 0xFF))
    tb(slide, icon, l + Inches(0.10), t + Inches(0.04),
       Inches(0.32), h - Inches(0.06),
       sz=title_sz, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
    tb(slide, title, l + Inches(0.50), t + Inches(0.04),
       w - Inches(0.58), h * 0.55,
       sz=title_sz, bold=True, color=WHITE)
    if sub:
        tb(slide, sub, l + Inches(0.50), t + h * 0.52,
           w - Inches(0.58), h * 0.42,
           sz=sub_sz, color=RGBColor(0xFF, 0xE8, 0xC0) if "orange" in str(fill) else RGBColor(0xB0, 0xCC, 0xF8))


# ━━━━━━━━━━━━━━━━━━━━━━━━━━ MAIN ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

prs = Presentation()
prs.slide_width  = SW
prs.slide_height = SH
blank = prs.slide_layouts[6]


# ╔══════════════════════════════════════════════════════════════════════╗
#  第 1 页  ·  机理分析（重绘）
#  ┌─────────────────────────────────────────────────────────────────┐
#  │  TITLE BAR                                                      │
#  ├─────────────────────────────────────────────────────────────────┤
#  │  全宽信号处理流程图（6步骤箭头流程）                             │
#  ├────────────────────────────┬────────────────────────────────────┤
#  │  【数据转换】 (橙色系)     │  【距离多普勒域信号级融合】(蓝色) │
#  │  子卡片①：双基地时延配准  │  子卡片①：RD图生成与信号模型      │
#  │  子卡片②：多普勒域统一    │  子卡片②：相参合并公式             │
#  │  子卡片③：坐标系配准      │  子卡片③：增益退化律               │
#  ├─────────────────────────────────────────────────────────────────┤
#  │  FOOTER                                                         │
#  └─────────────────────────────────────────────────────────────────┘
# ╚══════════════════════════════════════════════════════════════════════╝
s1 = prs.slides.add_slide(blank)
set_bg(s1, BG)
title_bar(s1,
          "空海平台多节点协同探测 · 相参合成机理分析",
          "Coherent Synthesis Mechanism: Range-Doppler Signal-Level Fusion & Data Conversion",
          1)

# ── 顶部：信号处理流程图 ────────────────────────────────────────────
FLOW_TOP = CT
FLOW_H   = Inches(0.80)
FLOW_PAD = Inches(0.06)

flow_steps = [
    ("原始回波\n(N=6链路)",    PIPE_COLS[0]),
    ("脉冲压缩\nPulse Comp",   PIPE_COLS[1]),
    ("距离多普勒图\nR-D Map",  PIPE_COLS[2]),
    ("数据转换\nData Conv",    PIPE_COLS[3]),   # 橙色
    ("信号级融合\nR-D Fusion", PIPE_COLS[4]),   # 深蓝
    ("CFAR检测\n目标输出",     PIPE_COLS[5]),
]
ARR_W  = Inches(0.14)
n_f    = len(flow_steps)
usable = CW - ARR_W * (n_f - 1)
step_w = usable / n_f
BOX_TOP = FLOW_TOP + Inches(0.16)
BOX_H   = FLOW_H - Inches(0.18)

px = MARGIN
for i, (lbl, col) in enumerate(flow_steps):
    rect(s1, px, BOX_TOP, step_w, BOX_H, col,
         lc=WHITE, lw=Pt(0.5), rnd=True)
    tb(s1, lbl, px, BOX_TOP, step_w, BOX_H,
       sz=8.5, bold=True, align=PP_ALIGN.CENTER, color=WHITE)
    if i < n_f - 1:
        tb(s1, "►", px + step_w, BOX_TOP + BOX_H * 0.25,
           ARR_W, BOX_H * 0.50,
           sz=10, color=COL_NAVY, align=PP_ALIGN.CENTER)
    px += step_w + ARR_W

# 在"数据转换"和"信号级融合"步骤下方加竖向虚线指示器
dt_x  = MARGIN + 3 * (step_w + ARR_W)
fus_x = MARGIN + 4 * (step_w + ARR_W)
for hx in [dt_x + step_w / 2, fus_x + step_w / 2]:
    line(s1, hx, BOX_TOP + BOX_H,
         hx, BOX_TOP + BOX_H + Inches(0.14),
         TITLE_BAR2, Pt(1.2))

# ── 主内容区 ────────────────────────────────────────────────────────
MAIN_TOP = FLOW_TOP + FLOW_H + Inches(0.14)
MAIN_H   = CB - MAIN_TOP
GAP_MID  = Inches(0.16)

# 左区：数据转换
L_W = (CW - GAP_MID) * 0.48
R_X = MARGIN + L_W + GAP_MID
R_W = CW - L_W - GAP_MID

# ── 左区大标题 ──
SEC_HDR_H = Inches(0.42)
# shadow
rect(s1, MARGIN + Inches(0.04), MAIN_TOP + Inches(0.04),
     L_W, SEC_HDR_H, RGBColor(0xCC, 0xD0, 0xCC))
rect(s1, MARGIN, MAIN_TOP, L_W, SEC_HDR_H, DT_HDR)
rect(s1, MARGIN, MAIN_TOP, Inches(0.10), SEC_HDR_H,
     RGBColor(0xFF, 0xD0, 0x80))  # bright accent left stripe
tb(s1, "⚙",
   MARGIN + Inches(0.14), MAIN_TOP + Inches(0.05),
   Inches(0.30), SEC_HDR_H - Inches(0.08),
   sz=14, bold=True, color=WHITE)
tb(s1, "数据转换   Data Conversion",
   MARGIN + Inches(0.50), MAIN_TOP + Inches(0.06),
   L_W - Inches(0.56), Inches(0.26),
   sz=13, bold=True, color=WHITE)
tb(s1, "双基地量测→统一参考域：时延配准 / 多普勒统一 / 坐标配准",
   MARGIN + Inches(0.50), MAIN_TOP + Inches(0.26),
   L_W - Inches(0.56), Inches(0.14),
   sz=8.5, color=RGBColor(0xFF, 0xE0, 0xB0))

# 左区三个子卡片
SUB_GAP  = Inches(0.10)
sub_top  = MAIN_TOP + SEC_HDR_H + Inches(0.10)
sub_h    = (MAIN_H - SEC_HDR_H - Inches(0.10) - 2 * SUB_GAP) / 3

dt_cards = [
    {
        "accent": DT_SUB1,
        "badge": "①",
        "title": "双基地时延配准（等效斜距折叠）",
        "items": [
            ("等效双程距离：r_eq = r_tx + r_rx  (双基地路径和)", False),
            ("r_fold = mod(r_eq, R_unamb)  ← 距离折叠处理",    True),
            ("rbin = round(r_fold / r_res) + 1  ← 距离门索引", True),
            ("精度要求：Δτ < λ/(4c) ≈ 0.5 ns ≡ Δr < 7.5 cm",  False),
        ]
    },
    {
        "accent": DT_SUB2,
        "badge": "②",
        "title": "多普勒域统一（速度转换）",
        "items": [
            ("双基地多普勒：f_d,k = (v_T·cosα_T + v_R·cosα_R) / λ", False),
            ("速度转换因子：K_v = cos(β/2)，β 为双基地角",           False),
            ("等效速度：v_eq = f_d·λ / (2·cos(β/2))",               True),
            ("各链路统一到等效单基地速度轴，消除多普勒偏差",         False),
        ]
    },
    {
        "accent": DT_SUB3,
        "badge": "③",
        "title": "坐标系配准与平台运动补偿",
        "items": [
            ("各平台坐标 → 统一 ECEF/ENU 参考系",               False),
            ("Δφ_motion = 2π·fc·Δτ_motion  ← 运动相位误差",     True),
            ("INS/DVL 辅助：精度 0.1 m/s、0.01°",               False),
            ("北斗/GPS 双频授时：时间基准同步 < 10 ns",          False),
        ]
    },
]

for ci, dc in enumerate(dt_cards):
    cy = sub_top + ci * (sub_h + SUB_GAP)
    sub_card(s1, MARGIN, cy, L_W, sub_h,
             dc["accent"], dc["title"], dc["items"],
             badge=dc["badge"])

# ── 右区大标题 ──
rect(s1, R_X + Inches(0.04), MAIN_TOP + Inches(0.04),
     R_W, SEC_HDR_H, RGBColor(0xC0, 0xC8, 0xE0))
rect(s1, R_X, MAIN_TOP, R_W, SEC_HDR_H, RD_HDR)
rect(s1, R_X, MAIN_TOP, Inches(0.10), SEC_HDR_H,
     RGBColor(0x80, 0xB8, 0xFF))
tb(s1, "⚡",
   R_X + Inches(0.14), MAIN_TOP + Inches(0.05),
   Inches(0.30), SEC_HDR_H - Inches(0.08),
   sz=14, bold=True, color=WHITE)
tb(s1, "距离多普勒域信号级融合",
   R_X + Inches(0.50), MAIN_TOP + Inches(0.06),
   R_W - Inches(0.56), Inches(0.26),
   sz=13, bold=True, color=WHITE)
tb(s1, "Range-Doppler Domain Signal-Level Coherent Fusion",
   R_X + Inches(0.50), MAIN_TOP + Inches(0.26),
   R_W - Inches(0.56), Inches(0.14),
   sz=8.5, color=RGBColor(0xB0, 0xD8, 0xFF))

# 右区三个子卡片
rd_cards = [
    {
        "accent": RD_SUB1,
        "badge": "①",
        "title": "RD图生成与信号模型",
        "items": [
            ("各链路距离多普勒图：X_k[r,d] = CFFT{ s_k(t)·e^{-j2πf_d·t} }",      False),
            ("sₖ(t) = A·exp[j(2π·fc·τₖ + φₖ)] + nₖ(t)  ← 复数基带信号",         True),
            ("信号级融合：在复数域（检测前）完成，保留完整相位信息",               False),
            ("区别于数据级（检测后）融合：信噪比增益 N² vs N",                     False),
        ]
    },
    {
        "accent": RD_SUB2,
        "badge": "②",
        "title": "相参合并公式（MRC加权）",
        "items": [
            ("X_coh[r,d] = Σ_k  w_k · X_k[r,d] · exp(j·φ_k)",                   True),
            ("MRC权值：w_k = SNR_k · coh_k · ph_cons_k  ← 三因子积",               False),
            ("相干度：coh_k = |Σ s_rot| / Σ|s_rot|  ← 回波自校准",                True),
            ("相位斜率一致性：ph_cons = exp(−|Δfd|/Δfd_ref)",                      True),
        ]
    },
    {
        "accent": RD_SUB3,
        "badge": "③",
        "title": "增益退化律与效能分析",
        "items": [
            ("理想相参增益：G_coh = 20·lg(N) = 15.6 dB（N=6）",                   False),
            ("非相参基准：G_incoh = 10·lg(N) = 7.8 dB，净增益 +7.8 dB",           False),
            ("含误差退化：G_eff = G · exp(−σ_φ²)",                                True),
            ("σ_φ < π/8 → G_eff ≥ 90%·G；σ_φ = π/4 → G_eff ≈ 64%·G",           False),
        ]
    },
]

for ci, rc in enumerate(rd_cards):
    cy = sub_top + ci * (sub_h + SUB_GAP)
    sub_card(s1, R_X, cy, R_W, sub_h,
             rc["accent"], rc["title"], rc["items"],
             badge=rc["badge"])

footer(s1,
       "核心机理：各链路双基地量测经数据转换配准至统一时延/多普勒/坐标域后，"
       "在距离多普勒复数域实施信号级相参合并，实现 N² 量级 SNR 增益")


# ╔══════════════════════════════════════════════════════════════════════╗
#  第 2 页  ·  相参合成方法设计（与 v5 保持一致）
# ╚══════════════════════════════════════════════════════════════════════╝
s2 = prs.slides.add_slide(blank)
set_bg(s2, BG)
title_bar(s2,
          "空海平台多节点协同探测 · 相参合成方法设计",
          "Coherent Synthesis Method Design — Processing Pipeline & Synchronization",
          2)

# ── 流水线 ─────────────────────────────────────────────────────────
PIPE_TOP = CT
PIPE_H   = Inches(0.78)

rect(s2, MARGIN, PIPE_TOP, CW * 0.50, PIPE_H,
     RGBColor(0xDC, 0xE8, 0xF8), lc=RGBColor(0xB0, 0xC8, 0xE8), lw=Pt(0.6))
rect(s2, MARGIN + CW * 0.50, PIPE_TOP, CW * 0.50, PIPE_H,
     RGBColor(0xD8, 0xEE, 0xE8), lc=RGBColor(0xA0, 0xCC, 0xBE), lw=Pt(0.6))
tb(s2, "Round 1 · 强目标检测（无MTI）",
   MARGIN + Inches(0.08), PIPE_TOP + Inches(0.02),
   CW * 0.48, Inches(0.18), sz=8, bold=True, color=COL_BLUE)
tb(s2, "Round 2 · 弱目标检测（三MTI并行）",
   MARGIN + CW * 0.50 + Inches(0.08), PIPE_TOP + Inches(0.02),
   CW * 0.48, Inches(0.18), sz=8, bold=True, color=COL_TEAL)
line(s2, MARGIN + CW * 0.50, PIPE_TOP, MARGIN + CW * 0.50, PIPE_TOP + PIPE_H,
     RGBColor(0xF5, 0xC5, 0x18), Pt(2.0))

pipe_steps = [
    ("总回波输入\nN=6链路",          PIPE_BLUE,   0.11),
    ("无MTI积累\nCFAR>12dB",         PIPE_BLUE,   0.12),
    ("GN定位\n强目标解算",            PIPE_PURPLE, 0.10),
    ("SIC\n强目标对消",               PIPE_RED,    0.09),
    ("三MTI并行\nSLOW/FAST/VFAST",   PIPE_TEAL,   0.13),
    ("CFAR>6dB\nNMS去重",             PIPE_TEAL,   0.11),
    ("RANSAC+GN\n精化定位",           PIPE_PURPLE, 0.11),
    ("卡尔曼跟踪\n相参增益估计",      PIPE_GREEN,  0.12),
    ("目标输出\n位置/速度/ID",        PIPE_GREEN,  0.11),
]
ARR_W2  = Inches(0.12)
n_s     = len(pipe_steps)
fracs   = [d[2] for d in pipe_steps]
arr_tot = ARR_W2 * (n_s - 1)
usable2 = CW - arr_tot
w_list  = [usable2 * f / sum(fracs) for f in fracs]

px2 = MARGIN
BOX_TOP2 = PIPE_TOP + Inches(0.20)
BOX_H2   = PIPE_H - Inches(0.22)
for i, (lbl, col, _) in enumerate(pipe_steps):
    sw = w_list[i]
    rect(s2, px2, BOX_TOP2, sw, BOX_H2, col,
         lc=WHITE, lw=Pt(0.5), rnd=True)
    tb(s2, lbl, px2, BOX_TOP2, sw, BOX_H2,
       sz=8, bold=True, align=PP_ALIGN.CENTER, color=WHITE)
    if i < n_s - 1:
        tb(s2, "►", px2 + sw, BOX_TOP2 + BOX_H2 * 0.28,
           ARR_W2, BOX_H2 * 0.44,
           sz=9, color=COL_NAVY, align=PP_ALIGN.CENTER)
    px2 += sw + ARR_W2

# ── 四列方法详解 ───────────────────────────────────────────────────
TBL_H   = Inches(1.50)
TBL_TOP = FOOTER_Y - TBL_H - Inches(0.42)
M4_TOP  = PIPE_TOP + PIPE_H + Inches(0.16)
M4_H    = TBL_TOP - M4_TOP - Inches(0.12)
GAP4    = Inches(0.12)
COL4W   = (CW - 3 * GAP4) / 4

HEADER_H_CARD = Inches(0.38)

method_sections = [
    {
        "num": "①", "color": COL_BLUE, "title": "时延对齐",
        "items": [
            ("北斗/GPS 双频授时 < 10 ns",      False),
            ("互相关峰值 + 亚像素插值",         False),
            ("r_fold = mod(r_eq, R_unamb)",     True),
            ("rbin = round(r_fold/r_res)+1",    True),
            ("INS/DVL 动态载体运动预测",        False),
            ("精度目标：Δτ < 0.5 ns",           False),
        ]
    },
    {
        "num": "②", "color": COL_TEAL, "title": "相位校正",
        "items": [
            ("相干度估计（cfar_detect_1d）",    False),
            ("s_rot = s·exp(−j2π·fd0·PRI·n)",  True),
            ("coh = |Σs_rot| / Σ|s_rot|",       True),
            ("相位斜率一致性（抑制镜像峰）",   False),
            ("ph_cons = exp(−|Δfd|/Δfd_ref)",   True),
            ("EKF/PLL 实时跟踪相位漂移",        False),
        ]
    },
    {
        "num": "③", "color": COL_PURPLE, "title": "MRC 合并",
        "items": [
            ("综合评分（cfar_detect_1d）",      False),
            ("score = SNR × (0.3+0.7·coh)",     True),
            ("       × (0.4+0.6·ph_cons)",       True),
            ("MRC 合并：Σ cell·e^{jφ}",        False),
            ("G = 10·lg(|coh|²/incoh)",         True),
            ("分层相参：质量异构时组内相参",    False),
        ]
    },
    {
        "num": "④", "color": COL_ORANGE, "title": "三MTI + SIC",
        "items": [
            ("三通道并行 MTI 滤波：",           False),
            ("SLOW  → [1,−1]    < 230 m/s",     True),
            ("FAST  → [1,−2,1]  230~400 m/s",   True),
            ("VFAST → [1,−3,3,−1]  > 400 m/s",  True),
            ("多帧积累：beta=0.65 powBank",      False),
            ("SIC 强目标对消 → 弱目标提取",     False),
        ]
    },
]

for ci, sec in enumerate(method_sections):
    cx = MARGIN + ci * (COL4W + GAP4)
    # Card shadow
    rect(s2, cx + Inches(0.04), M4_TOP + Inches(0.04), COL4W, M4_H,
         RGBColor(0xCC, 0xD8, 0xEC))
    # Card body
    rect(s2, cx, M4_TOP, COL4W, M4_H, CARD_BG,
         lc=CARD_BORDER, lw=Pt(0.8))
    # Header band
    rect(s2, cx, M4_TOP, COL4W, HEADER_H_CARD, sec["color"])
    # Badge
    circle_badge(s2,
                 cx + Inches(0.22), M4_TOP + HEADER_H_CARD / 2, 0.17,
                 WHITE, sec["num"], sz=9, tc=sec["color"])
    tb(s2, sec["title"],
       cx + Inches(0.50), M4_TOP + Inches(0.06),
       COL4W - Inches(0.56), HEADER_H_CARD - Inches(0.10),
       sz=12, bold=True, color=WHITE)
    line(s2, cx + Inches(0.08), M4_TOP + HEADER_H_CARD,
         cx + COL4W - Inches(0.06), M4_TOP + HEADER_H_CARD,
         RGBColor(0xE0, 0xEA, 0xF8), Pt(0.5))
    # Body
    body_top = M4_TOP + HEADER_H_CARD + Inches(0.02)
    body_h   = M4_H - HEADER_H_CARD
    rect(s2, cx + Inches(0.07), body_top,
         COL4W - Inches(0.07), body_h, RGBColor(0xF8, 0xFB, 0xFF))
    yy = body_top + Inches(0.07)
    item_h = Inches(0.29)
    for (txt, is_code) in sec["items"]:
        clr = RGBColor(0x1A, 0x6A, 0xC0) if is_code else DARK_TEXT
        prefix = "   " if is_code else "• "
        tb(s2, prefix + txt,
           cx + Inches(0.14), yy,
           COL4W - Inches(0.22), item_h,
           sz=9.2, color=clr)
        yy += item_h

# ── 性能指标表 ─────────────────────────────────────────────────────
tbl_label_y = TBL_TOP - Inches(0.36)
rect(s2, MARGIN, tbl_label_y, Inches(0.06), Inches(0.28), TITLE_BAR2)
tb(s2, "  关键性能指标 — 合并策略对比",
   MARGIN + Inches(0.10), tbl_label_y,
   Inches(8), Inches(0.28),
   sz=11, bold=True, color=COL_NAVY)
line(s2, MARGIN, tbl_label_y + Inches(0.30), MARGIN + CW,
     tbl_label_y + Inches(0.30), TITLE_BAR2, Pt(1.5))

tbl = s2.shapes.add_table(5, 5, MARGIN, TBL_TOP, CW, TBL_H).table


def cf(cell, text, sz=9.8, bold=False,
       bg=None, fg=DARK_TEXT, al=PP_ALIGN.CENTER):
    cell.text = text
    p = cell.text_frame.paragraphs[0]
    p.alignment = al
    r = p.runs[0] if p.runs else p.add_run()
    r.font.size = Pt(sz)
    r.font.bold = bold
    r.font.color.rgb = fg
    if bg:
        cell.fill.solid()
        cell.fill.fore_color.rgb = bg


headers = ["合并策略", "同步精度要求", "算法延迟", "MATLAB 参数 / 函数", "适用场景"]
for ci, h in enumerate(headers):
    cf(tbl.cell(0, ci), h, sz=10, bold=True, bg=TITLE_BAR, fg=WHITE)

tbl_data = [
    ["导频相参合并",    "时延 < 10 ns，Δφ < π/8",
     "< 1 ms",     "sys.fc=500M, cfar.min_snr_db=6",   "协同组网、慢机动"],
    ["回波自校准合并",  "时延 < 50 ns",
     "1~10 ms",    "track_aided_coherent, ≥3 链路",     "无导频、通信受限"],
    ["三MTI 分层合并",  "Δτ < 100 ns, coh ≥ 0.08",
     "低",          "mti_filter1/2/3, beta=0.65",        "节点多、同步异构"],
    ["SIC + 相参合并",  "强目标 SNR > 12 dB",
     "两轮 ~2 ms", "process_all_links_sic, sic_targets", "强弱目标共存"],
]
alt_bgs = [RGBColor(0xE8, 0xF2, 0xFF), CARD_BG,
           RGBColor(0xE8, 0xF2, 0xFF), CARD_BG]
for ri, row in enumerate(tbl_data):
    for ci, val in enumerate(row):
        cf(tbl.cell(ri + 1, ci), val, sz=9.5,
           bg=alt_bgs[ri], fg=DARK_TEXT,
           al=PP_ALIGN.LEFT if ci in (0, 3, 4) else PP_ALIGN.CENTER)

footer(s2,
       "核心权衡：空海高动态平台需在「估计精度 — 计算实时性 — 通信开销」三者间最优平衡；"
       "分层相参策略（SIC强目标 → 三MTI弱目标 → MRC合并）是算法设计核心")


# ── 保存 ───────────────────────────────────────────────────────────
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                   "coherent_synthesis_slides_v6.pptx")
prs.save(OUT)
print(f"✅  已生成：{OUT}")
