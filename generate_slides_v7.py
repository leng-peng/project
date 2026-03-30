"""
generate_slides_v7.py
─────────────────────────────────────────────────────────────────────────
多节点信号级高效相参合成 · 问题分析（v7）

  单页 PPT，布局与 v6 page-1 保持一致：
    ① 顶部全宽"分析思路"流程带（需求→科学问题→关键挑战→研究切入点）
    ② 主内容区 · 左列「需求→科学问题」（橙色系，3 个子卡片）
       - 科学问题 A：增益退化量化律
       - 科学问题 B：信号级融合的层次与增益跃升条件
       - 科学问题 C：相位估计可辨识性与 CRLB
    ③ 主内容区 · 右列「需要解决的关键问题」（蓝色系，3 个子卡片）
       - 关键问题①：异构双基地量测→统一域精度保持
       - 关键问题②：高动态时变相位建模与实时补偿
       - 关键问题③：自适应加权与实时性约束

运行：
    python generate_slides_v7.py
输出：
    coherent_synthesis_slides_v7.pptx
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
WHITE        = RGBColor(0xFF, 0xFF, 0xFF)
CARD_BORDER  = RGBColor(0xC8, 0xD8, 0xF0)
SEPARATOR    = RGBColor(0xC8, 0xD8, 0xF0)
FOOTER_BG    = RGBColor(0xE8, 0xF0, 0xFB)

# 主题色 — 科学问题（橙色系）
DT_HDR    = RGBColor(0xB5, 0x4A, 0x00)
DT_SUB1   = RGBColor(0xC8, 0x62, 0x10)
DT_SUB2   = RGBColor(0xD4, 0x78, 0x20)
DT_SUB3   = RGBColor(0xE0, 0x90, 0x30)

# 主题色 — 关键问题（蓝色系）
RD_HDR    = RGBColor(0x0D, 0x2A, 0x6E)
RD_SUB1   = RGBColor(0x1F, 0x6B, 0xC4)
RD_SUB2   = RGBColor(0x00, 0x7C, 0x78)
RD_SUB3   = RGBColor(0x6A, 0x2B, 0xB5)

# 流程带配色（4步）
FLOW_COLS = [
    RGBColor(0x1F, 0x6B, 0xC4),   # 工程需求
    RGBColor(0xB5, 0x4A, 0x00),   # 科学问题（橙）
    RGBColor(0x0D, 0x2A, 0x6E),   # 关键挑战（深蓝）
    RGBColor(0x00, 0x7C, 0x78),   # 研究切入点（青）
]

COL_NAVY   = RGBColor(0x0D, 0x2A, 0x6E)
FORMULA_BG = RGBColor(0xE8, 0xF2, 0xFF)

# ━━━━━━━━━━━━━━━━━━━━━━ 尺寸常量 ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
SW       = Inches(13.33)
SH       = Inches(7.50)
MARGIN   = Inches(0.32)
TITLE_H  = Inches(0.82)
CT       = TITLE_H + Inches(0.16)
FOOTER_Y = SH - Inches(0.50)
CB       = FOOTER_Y - Inches(0.04)
CW       = SW - 2 * MARGIN


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


# ─── 标题栏 ────────────────────────────────────────────────────────
def title_bar(slide, title_main, title_sub, page, total=1):
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


# ━━━━━━━━━━━━━━━━━━━━━━━━━━ MAIN ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

prs = Presentation()
prs.slide_width  = SW
prs.slide_height = SH
blank = prs.slide_layouts[6]

# ╔══════════════════════════════════════════════════════════════════════╗
#  问题分析页  ·  需求转化 + 关键问题点明
#  ┌─────────────────────────────────────────────────────────────────┐
#  │  TITLE BAR                                                      │
#  ├─────────────────────────────────────────────────────────────────┤
#  │  分析思路流程带（工程需求→科学问题→关键挑战→研究切入点）         │
#  ├──────────────────────────┬──────────────────────────────────────┤
#  │ 【需求→科学问题】(橙色)  │  【需要解决的关键问题】(蓝色)        │
#  │ ①增益退化量化律          │  ①异构量测→统一域精度保持           │
#  │ ②信号级融合层次与条件    │  ②时变相位建模与实时补偿             │
#  │ ③相位估计可辨识性        │  ③自适应加权与实时性约束             │
#  ├─────────────────────────────────────────────────────────────────┤
#  │  FOOTER                                                         │
#  └─────────────────────────────────────────────────────────────────┘
# ╚══════════════════════════════════════════════════════════════════════╝

s1 = prs.slides.add_slide(blank)
set_bg(s1, BG)
title_bar(s1,
          "多节点信号级高效相参合成 · 问题分析",
          "Problem Analysis: Requirement → Scientific Questions & Key Technical Challenges",
          1)

# ── 顶部：分析思路流程带 ─────────────────────────────────────────────
FLOW_TOP = CT
FLOW_H   = Inches(0.80)

flow_steps = [
    ("工程需求\n多节点协同高增益探测",    FLOW_COLS[0]),
    ("需求拆解\n量化→可研究科学问题",     FLOW_COLS[1]),
    ("科学问题\n增益律 / 融合条件 / CRLB", FLOW_COLS[2]),
    ("关键挑战\n量测配准 / 相位补偿 / 实时性", FLOW_COLS[3]),
]

ARR_W  = Inches(0.14)
n_f    = len(flow_steps)
usable = CW - ARR_W * (n_f - 1)
step_w = usable / n_f
BOX_TOP = FLOW_TOP + Inches(0.14)
BOX_H   = FLOW_H - Inches(0.16)

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

# 在"科学问题"和"关键挑战"步骤下方加竖向指示线
sci_x = MARGIN + 2 * (step_w + ARR_W)
key_x = MARGIN + 3 * (step_w + ARR_W)
for hx in [sci_x + step_w / 2, key_x + step_w / 2]:
    line(s1, hx, BOX_TOP + BOX_H,
         hx, BOX_TOP + BOX_H + Inches(0.14),
         TITLE_BAR2, Pt(1.2))

# ── 主内容区 ────────────────────────────────────────────────────────
MAIN_TOP = FLOW_TOP + FLOW_H + Inches(0.14)
MAIN_H   = CB - MAIN_TOP
GAP_MID  = Inches(0.16)

L_W = (CW - GAP_MID) * 0.48
R_X = MARGIN + L_W + GAP_MID
R_W = CW - L_W - GAP_MID

# ── 左区大标题（橙色：需求→科学问题）──
SEC_HDR_H = Inches(0.42)
rect(s1, MARGIN + Inches(0.04), MAIN_TOP + Inches(0.04),
     L_W, SEC_HDR_H, RGBColor(0xCC, 0xD0, 0xCC))
rect(s1, MARGIN, MAIN_TOP, L_W, SEC_HDR_H, DT_HDR)
rect(s1, MARGIN, MAIN_TOP, Inches(0.10), SEC_HDR_H,
     RGBColor(0xFF, 0xD0, 0x80))
tb(s1, "🔬",
   MARGIN + Inches(0.14), MAIN_TOP + Inches(0.05),
   Inches(0.30), SEC_HDR_H - Inches(0.08),
   sz=14, bold=True, color=WHITE)
tb(s1, "需求转化为可研究的科学问题",
   MARGIN + Inches(0.50), MAIN_TOP + Inches(0.06),
   L_W - Inches(0.56), Inches(0.26),
   sz=13, bold=True, color=WHITE)
tb(s1, "将工程指标映射为可量化、可推导、可验证的科学命题",
   MARGIN + Inches(0.50), MAIN_TOP + Inches(0.26),
   L_W - Inches(0.56), Inches(0.14),
   sz=8.5, color=RGBColor(0xFF, 0xE0, 0xB0))

# 左区三个子卡片
SUB_GAP = Inches(0.10)
sub_top = MAIN_TOP + SEC_HDR_H + Inches(0.10)
sub_h   = (MAIN_H - SEC_HDR_H - Inches(0.10) - 2 * SUB_GAP) / 3

sci_cards = [
    {
        "accent": DT_SUB1,
        "badge": "A",
        "title": "增益退化量化律（需求：高增益→量化关系）",
        "items": [
            ('工程需求"高增益"→ 问：G 如何随同步误差 (σ_τ, σ_φ, σ_f) 退化？', False),
            ("G_eff = G · exp(−σ_φ²)，G = 20·lg(N) = 15.6 dB（N=6）",         True),
            ("效能临界点：σ_φ < π/8 → G_eff ≥ 90%·G；超出则急剧退化",         False),
            ("目标：建立误差源→合并增益的端到端量化传播模型",                   False),
        ]
    },
    {
        "accent": DT_SUB2,
        "badge": "B",
        "title": "信号级融合层次与增益跃升条件（vs 数据级融合）",
        "items": [
            ("需求：信号级融合 → 问：O(N²) vs O(N) 增益跃升的充要条件？",       False),
            ("检测前复数域融合保留完整相位 → SNR增益 N²；检测后丢相位 → N",    False),
            ("双基地量测 (r_eq, f_d,bist) → 统一R-D域，映射精度保持是前提",    False),
            ("r_fold = mod(r_eq, R_unamb)，K_v = cos(β/2) 等效折叠精度？",     True),
        ]
    },
    {
        "accent": DT_SUB3,
        "badge": "C",
        "title": "相位估计可辨识性与 CRLB（无导频自校准）",
        "items": [
            ("需求：无导频回波自校准 → 问：相位偏差估计 CRLB 由什么决定？",     False),
            ("coh_k = |Σ s_rot| / Σ|s_rot|，CRLB ∝ 1/(N_c · SNR_k)",         True),
            ("高动态时变相位：Δφ_k(t) = 2π·fc·Δτ_motion(t)，非平稳建模难题", False),
            ("问：需要多少节点才能保证无导频估计可收敛？",                       False),
        ]
    },
]

for ci, dc in enumerate(sci_cards):
    cy = sub_top + ci * (sub_h + SUB_GAP)
    sub_card(s1, MARGIN, cy, L_W, sub_h,
             dc["accent"], dc["title"], dc["items"],
             badge=dc["badge"])

# ── 右区大标题（蓝色：关键问题）──
rect(s1, R_X + Inches(0.04), MAIN_TOP + Inches(0.04),
     R_W, SEC_HDR_H, RGBColor(0xC0, 0xC8, 0xE0))
rect(s1, R_X, MAIN_TOP, R_W, SEC_HDR_H, RD_HDR)
rect(s1, R_X, MAIN_TOP, Inches(0.10), SEC_HDR_H,
     RGBColor(0x80, 0xB8, 0xFF))
tb(s1, "🔑",
   R_X + Inches(0.14), MAIN_TOP + Inches(0.05),
   Inches(0.30), SEC_HDR_H - Inches(0.08),
   sz=14, bold=True, color=WHITE)
tb(s1, "需要解决的关键问题",
   R_X + Inches(0.50), MAIN_TOP + Inches(0.06),
   R_W - Inches(0.56), Inches(0.26),
   sz=13, bold=True, color=WHITE)
tb(s1, "Key Technical Challenges — Bottlenecks Blocking Practical Deployment",
   R_X + Inches(0.50), MAIN_TOP + Inches(0.26),
   R_W - Inches(0.56), Inches(0.14),
   sz=8.5, color=RGBColor(0xB0, 0xD8, 0xFF))

# 右区三个子卡片
key_cards = [
    {
        "accent": RD_SUB1,
        "badge": "①",
        "title": "异构双基地量测→统一域转换的精度保持",
        "items": [
            ("核心矛盾：各节点双基地角 β_k 不同，r_eq 与 f_d 无法直接对齐",   False),
            ("r_fold = mod(r_eq, R_unamb) 在大双基地角时距离门是否错位？",     True),
            ("K_v = cos(β/2) 在 β > 60° 时引入的速度误差有多大？",            True),
            ("精度损失与相参增益收益之间的权衡无法简单公式化，需数值求解",     False),
        ]
    },
    {
        "accent": RD_SUB2,
        "badge": "②",
        "title": "高动态时变相位误差的建模与实时补偿",
        "items": [
            ("Δφ_k(t) = 2π·fc·Δτ_k(t) + δφ_osc,k(t)  ← 持续时变，非恒定", True),
            ("GPS/北斗授时精度 10~50 ns vs 相参要求 Δτ < 0.5 ns，差 1~2 量级", False),
            ("互相关亚像素插值能否弥补？ → 低SNR时估计误差反而加剧退化",      False),
            ("待解决：σ_est—增益G—帧长N_c 的联合最优设计准则",               False),
        ]
    },
    {
        "accent": RD_SUB3,
        "badge": "③",
        "title": "质量异构节点自适应加权与计算实时性",
        "items": [
            ('w_k = SNR_k · coh_k · ph_cons_k，等权合并使劣质节点"污染"输出', False),
            ("相位估计失败时须自动识别降权，避免合并输出突变（稳健机制）",     False),
            ("计算量 O(N·N_c·N_sc·lgN_c) ≈ 2.4×10⁶ 次/帧，需在 9 ms 内完成",True),
            ("近似算法（稀疏采样/子带分解）：精度损失 vs 复杂度降一量级权衡", False),
        ]
    },
]

for ci, kc in enumerate(key_cards):
    cy = sub_top + ci * (sub_h + SUB_GAP)
    sub_card(s1, R_X, cy, R_W, sub_h,
             kc["accent"], kc["title"], kc["items"],
             badge=kc["badge"])

footer(s1,
       "核心纽带：「残余相位误差 σ_φ」连接所有关键问题 —— "
       "量测转换精度 + 平台运动补偿 + 授时精度三者共同决定 σ_φ 分布，"
       "进而决定信号级融合能否突破 O(N²) 增益上界")


# ── 保存 ───────────────────────────────────────────────────────────
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                   "coherent_synthesis_slides_v7.pptx")
prs.save(OUT)
print(f"✅  已生成：{OUT}")
