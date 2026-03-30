"""
generate_slides_v5.py
─────────────────────────────────────────────────────────────────────────
空海平台多节点协同探测 — 相参合成两页 PPT（学术浅色模板 v5）

  Page 1：相参合成 · 机理分析
  Page 2：相参合成 · 方法设计

模板风格（v5）：
  • 浅色背景（白色 / 极浅蓝灰），区别于 v4 深海蓝夜间主题
  • 顶部标题栏：深蓝渐变色（#1F3864→#2E74B5）+ 白色标题
  • 每个内容块：白色卡片 + 彩色顶部标题带 + 阴影描边
  • 序号徽章：圆形彩色背景
  • 流水线：水平圆角色块 + 方向箭头
  • 页脚：浅灰分割线 + 蓝色要点文字

运行：
    python generate_slides_v5.py
输出：
    coherent_synthesis_slides_v5.pptx
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# ━━━━━━━━━━━━━━━━━━━━━━ 调色板 ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
BG           = RGBColor(0xF4, 0xF7, 0xFC)   # 极浅蓝灰背景
CARD_BG      = RGBColor(0xFF, 0xFF, 0xFF)   # 卡片白色背景
TITLE_BAR    = RGBColor(0x1F, 0x38, 0x64)   # 深海蓝标题栏
TITLE_BAR2   = RGBColor(0x2E, 0x74, 0xB5)   # 浅蓝（用于装饰）
DARK_TEXT    = RGBColor(0x1A, 0x1A, 0x3A)   # 深色正文
MID_TEXT     = RGBColor(0x3D, 0x3D, 0x6B)   # 中深色辅助文字
LIGHT_TEXT   = RGBColor(0x6E, 0x8E, 0xBF)   # 浅蓝灰提示文字
WHITE        = RGBColor(0xFF, 0xFF, 0xFF)
CARD_BORDER  = RGBColor(0xC8, 0xD8, 0xF0)   # 卡片描边浅蓝
SEPARATOR    = RGBColor(0xC8, 0xD8, 0xF0)   # 分割线
FOOTER_BG    = RGBColor(0xE8, 0xF0, 0xFB)   # 页脚背景

# 功能性色彩 — 卡片标题带
COL_BLUE     = RGBColor(0x1F, 0x6B, 0xC4)   # 蓝
COL_TEAL     = RGBColor(0x00, 0x7C, 0x78)   # 青
COL_PURPLE   = RGBColor(0x6A, 0x2B, 0xB5)   # 紫
COL_ORANGE   = RGBColor(0xC8, 0x55, 0x0A)   # 橙
COL_GREEN    = RGBColor(0x15, 0x7A, 0x35)   # 绿
COL_RED      = RGBColor(0xA8, 0x12, 0x12)   # 红
COL_NAVY     = RGBColor(0x0D, 0x2A, 0x6E)   # 深蓝导航色

# 流水线模块色（统一浅色系）
PIPE_BLUE    = RGBColor(0x1F, 0x6B, 0xC4)
PIPE_RED     = RGBColor(0xA8, 0x12, 0x12)
PIPE_TEAL    = RGBColor(0x00, 0x7C, 0x78)
PIPE_PURPLE  = RGBColor(0x6A, 0x2B, 0xB5)
PIPE_GREEN   = RGBColor(0x15, 0x7A, 0x35)

# ━━━━━━━━━━━━━━━━━━━━━━ 尺寸常量 ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
SW       = Inches(13.33)
SH       = Inches(7.50)
MARGIN   = Inches(0.32)
TITLE_H  = Inches(0.82)
CT       = TITLE_H + Inches(0.18)    # content top
FOOTER_Y = SH - Inches(0.50)        # footer top
CB       = FOOTER_Y - Inches(0.05)  # content bottom
CW       = SW - 2 * MARGIN
COL_H    = CB - CT


# ━━━━━━━━━━━━━━━━━━━━━━ 工具函数 ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def set_bg(slide, color):
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = color


def rect(slide, l, t, w, h, fill,
         lc=None, lw=Pt(0.75), rnd=False):
    s = slide.shapes.add_shape(5 if rnd else 1, l, t, w, h)
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    if lc:
        s.line.color.rgb = lc
        s.line.width = lw
    else:
        s.line.fill.background()
    return s


def line(slide, x1, y1, x2, y2=None,
         color=SEPARATOR, width=Pt(0.6)):
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


def circle_badge(slide, cx, cy, r_inch, fill, text,
                 sz=11, bold=True, tc=WHITE):
    d = Inches(r_inch * 2)
    s = slide.shapes.add_shape(9,
                               cx - Inches(r_inch),
                               cy - Inches(r_inch), d, d)
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


# ─── 标题栏 ────────────────────────────────────────────────────────
def title_bar(slide, title_main, title_sub, page, total=2):
    """Two-tone title bar: dark navy left band + blue body."""
    # Left accent block
    rect(slide, Inches(0), Inches(0), Inches(0.18), TITLE_H,
         COL_NAVY)
    # Main title bar
    rect(slide, Inches(0.18), Inches(0), SW - Inches(0.18), TITLE_H,
         TITLE_BAR)
    # Thin bottom accent line
    line(slide, Inches(0), TITLE_H, SW, TITLE_H,
         TITLE_BAR2, Pt(2.5))
    # Title main text
    tb(slide, title_main,
       Inches(0.36), Inches(0.10), Inches(10.0), Inches(0.38),
       sz=20, bold=True, color=WHITE)
    # Title sub text
    tb(slide, title_sub,
       Inches(0.36), Inches(0.50), Inches(10.0), Inches(0.26),
       sz=10.5, bold=False, color=RGBColor(0xB0, 0xCC, 0xF0))
    # Page badge
    badge_cx = SW - Inches(0.52)
    badge_cy = TITLE_H / 2
    circle_badge(slide, badge_cx, badge_cy, 0.28,
                 TITLE_BAR2, f"{page}/{total}", sz=11)


# ─── 页脚 ──────────────────────────────────────────────────────────
def footer(slide, text):
    rect(slide, Inches(0), FOOTER_Y, SW, SH - FOOTER_Y, FOOTER_BG)
    line(slide, MARGIN, FOOTER_Y, SW - MARGIN, FOOTER_Y,
         TITLE_BAR2, Pt(1.5))
    rect(slide, MARGIN, FOOTER_Y + Inches(0.10),
         Inches(0.05), Inches(0.28), TITLE_BAR2)
    tb(slide, "  " + text,
       MARGIN + Inches(0.10), FOOTER_Y + Inches(0.09),
       CW - Inches(0.14), Inches(0.34),
       sz=8.8, italic=True, color=COL_NAVY)


# ─── 内容卡片 ──────────────────────────────────────────────────────
def card(slide, l, t, w, h, accent, header_text, body_lines,
         badge_text=None, header_sz=11.5, body_sz=9.5):
    """
    White card with colored top band.
    body_lines: list of (text, is_formula)
    """
    HEADER_H = Inches(0.38)
    # Card shadow simulation (dark rect slightly offset)
    rect(slide, l + Inches(0.04), t + Inches(0.04), w, h,
         RGBColor(0xCC, 0xD8, 0xEC))
    # Card body
    rect(slide, l, t, w, h, CARD_BG,
         lc=CARD_BORDER, lw=Pt(0.8))
    # Colored header band
    rect(slide, l, t, w, HEADER_H, accent)
    # Badge
    bx = l + Inches(0.22)
    if badge_text:
        circle_badge(slide, bx, t + HEADER_H / 2, 0.17,
                     RGBColor(0xFF, 0xFF, 0xFF), badge_text,
                     sz=9, bold=True, tc=accent)
    # Header text
    hx = l + (Inches(0.50) if badge_text else Inches(0.14))
    tb(slide, header_text, hx, t + Inches(0.06),
       w - hx + l - Inches(0.08), HEADER_H - Inches(0.10),
       sz=header_sz, bold=True, color=WHITE)
    # Thin white sep under header
    line(slide, l + Inches(0.08), t + HEADER_H,
         l + w - Inches(0.08), t + HEADER_H,
         RGBColor(0xE0, 0xEA, 0xF8), Pt(0.5))
    # Body items
    yy = t + HEADER_H + Inches(0.07)
    item_h = body_sz * 1.75 * 914.4 / 72   # approx line height in EMUs
    for (txt, is_formula) in body_lines:
        clr = RGBColor(0x1A, 0x6A, 0xC0) if is_formula else DARK_TEXT
        prefix = "   " if is_formula else "• "
        tb(slide, prefix + txt,
           l + Inches(0.12), yy,
           w - Inches(0.18), item_h,
           sz=body_sz, color=clr)
        yy += item_h + Inches(0.01)


# ─── 公式高亮块 ────────────────────────────────────────────────────
def formula_box(slide, l, t, w, h, text, sz=10.5):
    rect(slide, l, t, w, h,
         RGBColor(0xE8, 0xF2, 0xFF),
         lc=TITLE_BAR2, lw=Pt(0.8))
    tb(slide, text, l + Inches(0.10), t + Inches(0.06),
       w - Inches(0.18), h - Inches(0.10),
       sz=sz, color=RGBColor(0x0D, 0x2A, 0x6E), bold=True)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━ MAIN ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

prs  = Presentation()
prs.slide_width  = SW
prs.slide_height = SH
blank = prs.slide_layouts[6]   # completely blank layout


# ╔══════════════════════════════════════════════════════════════════════╗
#  第 1 页  ·  相参合成机理分析
#  布局（3列）：
#    左列(52%)：4 个同步误差卡片（2×2）
#    右列(46%)：上→信号模型 | 中→增益退化律 | 下→CRLB+系统参数
# ╚══════════════════════════════════════════════════════════════════════╝
s1 = prs.slides.add_slide(blank)
set_bg(s1, BG)
title_bar(s1,
          "空海平台多节点协同探测 · 相参合成机理分析",
          "Coherent Synthesis Mechanism — Air-Sea Multi-Node Distributed Radar",
          1)

# ── 左区：4 个误差卡片（2×2 网格） ──────────────────────────────────
L_COL_W  = CW * 0.52
R_COL_X  = MARGIN + L_COL_W + Inches(0.14)
R_COL_W  = CW - L_COL_W - Inches(0.14)

GAP   = Inches(0.10)
CARD_W = (L_COL_W - GAP) / 2
CARD_H = (COL_H - GAP) / 2

errors = [
    {
        "num": "①", "color": COL_BLUE, "title": "时延同步误差",
        "body": [
            ("几何双程时延差 → 脉冲错位",            False),
            ("精度要求：Δτ < λ/(4c) ≈ 0.5 ns",      False),
            ("r_fold = mod(r_eq, R_unamb)",           True),
            ("tau = 2·r_fold / c",                    True),
            ("方法：互相关 + 北斗/GPS 双频授时",     False),
        ]
    },
    {
        "num": "②", "color": COL_TEAL, "title": "相位同步误差",
        "body": [
            ("载频相位不一致 → 复数旋转偏差",        False),
            ("cell_c = cellv·exp(j2π·fc·τ)",         True),
            ("coh_sum += cell_c",                     True),
            ("方法：导频估计 + EKF/PLL 联合跟踪",   False),
            ("无导频：回波自校准",                   False),
        ]
    },
    {
        "num": "③", "color": COL_PURPLE, "title": "频率同步误差",
        "body": [
            ("多普勒 + 振荡器频偏 → 相位漂移",      False),
            ("ΔΦ = 2π · Δf · t（线性漂移）",        False),
            ("ph_cons = exp(−|Δfd|/Δfd_ref)",        True),
            ("方法：相位斜率一致性检验",             False),
            ("PLL 实时跟踪频偏",                     False),
        ]
    },
    {
        "num": "④", "color": COL_ORANGE, "title": "幅度一致性",
        "body": [
            ("通道增益不均 → 合并旁瓣抬高",         False),
            ("score = SNR × (0.3+0.7·coh)",          True),
            ("      × (0.4+0.6·ph_cons)",             True),
            ("方法：MRC 自适应加权合并",             False),
            ("质量低链路自动降权",                   False),
        ]
    },
]

for i, err in enumerate(errors):
    col = i % 2
    row = i // 2
    ex = MARGIN + col * (CARD_W + GAP)
    ey = CT + row * (CARD_H + GAP)
    card(s1, ex, ey, CARD_W, CARD_H,
         err["color"], err["title"], err["body"],
         badge_text=err["num"])

# ── 右区：三个竖向段 ──────────────────────────────────────────────
# 信号模型块（上段，约 38%）
SEG1_H = COL_H * 0.36
rect(s1, R_COL_X - Inches(0.04), CT + Inches(0.04), R_COL_W, SEG1_H,
     RGBColor(0xCC, 0xD8, 0xEC))  # shadow
rect(s1, R_COL_X, CT, R_COL_W, SEG1_H, CARD_BG,
     lc=CARD_BORDER, lw=Pt(0.8))
rect(s1, R_COL_X, CT, R_COL_W, Inches(0.38), COL_BLUE)
tb(s1, "信号模型（OFDM 双基地体制）",
   R_COL_X + Inches(0.12), CT + Inches(0.06),
   R_COL_W - Inches(0.18), Inches(0.28),
   sz=11.5, bold=True, color=WHITE)

tb(s1, "系统：2 Tx（海面 15~18 m/s）+ 3 Rx（空中 3500~3800 m）= 6 双基地链路",
   R_COL_X + Inches(0.12), CT + Inches(0.42),
   R_COL_W - Inches(0.18), Inches(0.22),
   sz=9.5, color=MID_TEXT)
formula_box(
    s1,
    R_COL_X + Inches(0.10), CT + Inches(0.68),
    R_COL_W - Inches(0.18), Inches(0.38),
    "sₖ(t) = A · exp[ j(2π·fc·τₖ + φₖ) ] + nₖ(t)"
)
tb(s1,
   "coh_sum = Σₖ cellₖ · exp(j·2π·fc·τₖ)\n"
   "SNR_coh ∝ N²     vs.     SNR_incoh ∝ N",
   R_COL_X + Inches(0.12), CT + Inches(1.10),
   R_COL_W - Inches(0.18), Inches(0.36),
   sz=10, color=RGBColor(0x0D, 0x2A, 0x6E), bold=True)

# 增益退化律（中段）
SEG2_TOP = CT + SEG1_H + Inches(0.12)
SEG2_H   = COL_H * 0.33
rect(s1, R_COL_X - Inches(0.04), SEG2_TOP + Inches(0.04), R_COL_W, SEG2_H,
     RGBColor(0xCC, 0xD8, 0xEC))
rect(s1, R_COL_X, SEG2_TOP, R_COL_W, SEG2_H, CARD_BG,
     lc=CARD_BORDER, lw=Pt(0.8))
rect(s1, R_COL_X, SEG2_TOP, R_COL_W, Inches(0.38), COL_TEAL)
tb(s1, "相参增益退化律",
   R_COL_X + Inches(0.12), SEG2_TOP + Inches(0.06),
   R_COL_W - Inches(0.18), Inches(0.28),
   sz=11.5, bold=True, color=WHITE)

formula_box(
    s1,
    R_COL_X + Inches(0.10), SEG2_TOP + Inches(0.42),
    R_COL_W - Inches(0.18), Inches(0.32),
    "G = 10·lg( |Σ cellₖ·e^{j2πfcτₖ}|² / Σ|cellₖ|² )"
)
gain_txt = (
    "σ²_φ = σ²_time·(2πfc)² + σ²_phase + σ²_freq·T²\n"
    "σ_φ < π/8  →  G > 0.90·N²（效能 ≥ 90%）\n"
    "σ_φ = π/4  →  G ≈ 0.64·N²（退化 36%）"
)
tb(s1, gain_txt,
   R_COL_X + Inches(0.12), SEG2_TOP + Inches(0.78),
   R_COL_W - Inches(0.18), Inches(0.50),
   sz=9.8, color=DARK_TEXT)

# CRLB + 系统参数（下段）
SEG3_TOP = SEG2_TOP + SEG2_H + Inches(0.12)
SEG3_H   = CB - SEG3_TOP
rect(s1, R_COL_X - Inches(0.04), SEG3_TOP + Inches(0.04), R_COL_W, SEG3_H,
     RGBColor(0xCC, 0xD8, 0xEC))
rect(s1, R_COL_X, SEG3_TOP, R_COL_W, SEG3_H, CARD_BG,
     lc=CARD_BORDER, lw=Pt(0.8))
rect(s1, R_COL_X, SEG3_TOP, R_COL_W, Inches(0.38), COL_PURPLE)
tb(s1, "CRLB 下界 & 系统参数",
   R_COL_X + Inches(0.12), SEG3_TOP + Inches(0.06),
   R_COL_W - Inches(0.18), Inches(0.28),
   sz=11.5, bold=True, color=WHITE)

crlb_txt = (
    "测距下界：σ_r = c / (2B√(2·SNR))\n"
    "测速下界：σ_v = λ / (2Tc√SNR)\n"
    "FIM：J = Σₖ HₖᵀRₖ⁻¹Hₖ  →  6链路 ≈ 100 m\n"
)
tb(s1, crlb_txt,
   R_COL_X + Inches(0.12), SEG3_TOP + Inches(0.42),
   R_COL_W - Inches(0.18), Inches(0.58),
   sz=9.5, color=DARK_TEXT)

# Compact params table
params = [
    ("fc=500 MHz", "B=2 MHz"),
    ("Nsc=1024",   "PRI=140 μs"),
    ("Nc=64脉冲",  "r_res=75 m"),
    ("N=6链路",    "λ=0.6 m"),
]
row_h   = Inches(0.24)
tab_top = SEG3_TOP + Inches(1.06)
for j, (a, b) in enumerate(params):
    ry = tab_top + j * row_h
    bg = RGBColor(0xE8, 0xF2, 0xFF) if j % 2 == 0 else CARD_BG
    rect(s1, R_COL_X + Inches(0.08), ry,
         R_COL_W - Inches(0.14), row_h, bg,
         lc=CARD_BORDER, lw=Pt(0.4))
    half_w = (R_COL_W - Inches(0.14)) / 2
    tb(s1, a, R_COL_X + Inches(0.12), ry + Inches(0.03),
       half_w - Inches(0.06), row_h - Inches(0.04),
       sz=9, bold=True, color=COL_NAVY)
    tb(s1, b, R_COL_X + Inches(0.12) + half_w, ry + Inches(0.03),
       half_w - Inches(0.06), row_h - Inches(0.04),
       sz=9, bold=True, color=COL_NAVY)

footer(s1,
       "核心挑战：空海高动态平台四大误差耦合非线性退化 — "
       "同步要求 Δτ < 0.5 ns、Δφ < π/8（47 mrad），远超地基组网系统")


# ╔══════════════════════════════════════════════════════════════════════╗
#  第 2 页  ·  相参合成方法设计
#  布局：
#    顶部（全宽）：双轮处理流水线
#    中部（4等列）：方法块（时延对齐 / 相位校正 / MRC合并 / 三MTI+SIC）
#    底部（全宽）：关键性能指标表
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

# 分组背景
r1_end = CW * 0.50
rect(s2, MARGIN, PIPE_TOP, CW * 0.50, PIPE_H,
     RGBColor(0xDC, 0xE8, 0xF8),
     lc=RGBColor(0xB0, 0xC8, 0xE8), lw=Pt(0.6))
rect(s2, MARGIN + CW * 0.50, PIPE_TOP, CW * 0.50, PIPE_H,
     RGBColor(0xD8, 0xEE, 0xE8),
     lc=RGBColor(0xA0, 0xCC, 0xBE), lw=Pt(0.6))
tb(s2, "Round 1 · 强目标检测（无MTI）",
   MARGIN + Inches(0.08), PIPE_TOP + Inches(0.02),
   CW * 0.48, Inches(0.18),
   sz=8, bold=True, color=COL_BLUE)
tb(s2, "Round 2 · 弱目标检测（三MTI并行）",
   MARGIN + CW * 0.50 + Inches(0.08), PIPE_TOP + Inches(0.02),
   CW * 0.48, Inches(0.18),
   sz=8, bold=True, color=COL_TEAL)
# Gold divider between R1/R2
line(s2,
     MARGIN + CW * 0.50, PIPE_TOP,
     MARGIN + CW * 0.50, PIPE_TOP + PIPE_H,
     RGBColor(0xF5, 0xC5, 0x18), Pt(2.0))

pipe_steps = [
    ("总回波输入\nN=6链路",           PIPE_BLUE,   0.11),
    ("无MTI积累\nCFAR>12dB",          PIPE_BLUE,   0.12),
    ("GN定位\n强目标解算",             PIPE_PURPLE, 0.10),
    ("SIC\n强目标对消",                PIPE_RED,    0.09),
    ("三MTI并行\nSLOW/FAST/VFAST",    PIPE_TEAL,   0.13),
    ("CFAR>6dB\nNMS去重",              PIPE_TEAL,   0.11),
    ("RANSAC+GN\n精化定位",            PIPE_PURPLE, 0.11),
    ("卡尔曼跟踪\n相参增益估计",       PIPE_GREEN,  0.12),
    ("目标输出\n位置/速度/ID",         PIPE_GREEN,  0.11),
]
ARR_W   = Inches(0.12)
n_s     = len(pipe_steps)
fracs   = [d[2] for d in pipe_steps]
arr_tot = ARR_W * (n_s - 1)
usable  = CW - arr_tot
w_list  = [usable * f / sum(fracs) for f in fracs]

px = MARGIN
BOX_TOP = PIPE_TOP + Inches(0.20)
BOX_H   = PIPE_H   - Inches(0.22)
for i, (lbl, col, _) in enumerate(pipe_steps):
    sw = w_list[i]
    rect(s2, px, BOX_TOP, sw, BOX_H, col,
         lc=WHITE, lw=Pt(0.5), rnd=True)
    tb(s2, lbl, px, BOX_TOP, sw, BOX_H,
       sz=8, bold=True, align=PP_ALIGN.CENTER, color=WHITE)
    if i < n_s - 1:
        tb(s2, "►", px + sw, BOX_TOP + BOX_H * 0.28,
           ARR_W, BOX_H * 0.44,
           sz=9, color=COL_NAVY, align=PP_ALIGN.CENTER)
    px += sw + ARR_W

# ── 四列方法详解 ───────────────────────────────────────────────────
TBL_H   = Inches(1.50)
TBL_TOP = FOOTER_Y - TBL_H - Inches(0.42)
M4_TOP  = PIPE_TOP + PIPE_H + Inches(0.16)
M4_H    = TBL_TOP - M4_TOP - Inches(0.12)
GAP4    = Inches(0.12)
COL4W   = (CW - 3 * GAP4) / 4

method_sections = [
    {
        "num": "①", "color": COL_BLUE, "title": "时延对齐",
        "items": [
            ("北斗/GPS 双频授时 < 10 ns",          False),
            ("互相关峰值 + 亚像素插值",             False),
            ("r_fold = mod(r_eq, R_unamb)",         True),
            ("rbin = round(r_fold/r_res)+1",        True),
            ("INS/DVL 动态载体运动预测",            False),
            ("精度目标：Δτ < 0.5 ns",               False),
        ]
    },
    {
        "num": "②", "color": COL_TEAL, "title": "相位校正",
        "items": [
            ("相干度估计（cfar_detect_1d）",        False),
            ("s_rot = s·exp(−j2π·fd0·PRI·n)",      True),
            ("coh = |Σs_rot| / Σ|s_rot|",           True),
            ("相位斜率一致性（抑制镜像峰）",       False),
            ("ph_cons = exp(−|Δfd|/Δfd_ref)",       True),
            ("EKF/PLL 实时跟踪相位漂移",            False),
        ]
    },
    {
        "num": "③", "color": COL_PURPLE, "title": "MRC 合并",
        "items": [
            ("综合评分（cfar_detect_1d）",          False),
            ("score = SNR × (0.3+0.7·coh)",         True),
            ("       × (0.4+0.6·ph_cons)",           True),
            ("MRC 合并：Σ cell·e^{jφ}",            False),
            ("G = 10·lg(|coh|²/incoh)",             True),
            ("分层相参：质量异构时组内相参",        False),
        ]
    },
    {
        "num": "④", "color": COL_ORANGE, "title": "三MTI + SIC",
        "items": [
            ("三通道并行 MTI 滤波：",               False),
            ("SLOW  → [1,−1]    < 230 m/s",         True),
            ("FAST  → [1,−2,1]  230~400 m/s",       True),
            ("VFAST → [1,−3,3,−1]  > 400 m/s",      True),
            ("多帧积累：beta=0.65 powBank",          False),
            ("SIC 强目标对消 → 弱目标提取",         False),
        ]
    },
]

for ci, sec in enumerate(method_sections):
    cx = MARGIN + ci * (COL4W + GAP4)
    card(s2, cx, M4_TOP, COL4W, M4_H,
         sec["color"], sec["title"], sec["items"],
         badge_text=sec["num"])

# ── 性能指标表 ─────────────────────────────────────────────────────
tbl_label_y = TBL_TOP - Inches(0.36)
rect(s2, MARGIN, tbl_label_y,
     Inches(0.06), Inches(0.28), TITLE_BAR2)
tb(s2, "  关键性能指标 — 合并策略对比",
   MARGIN + Inches(0.10), tbl_label_y,
   Inches(8), Inches(0.28),
   sz=11, bold=True, color=COL_NAVY)
line(s2, MARGIN, tbl_label_y + Inches(0.30),
     MARGIN + CW, tbl_label_y + Inches(0.30),
     TITLE_BAR2, Pt(1.5))

tbl = s2.shapes.add_table(5, 5,
                           MARGIN, TBL_TOP,
                           CW, TBL_H).table


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


HDR_BG = TITLE_BAR
headers = ["合并策略", "同步精度要求", "算法延迟", "MATLAB 参数 / 函数", "适用场景"]
for ci, h in enumerate(headers):
    cf(tbl.cell(0, ci), h, sz=10, bold=True, bg=HDR_BG, fg=WHITE)

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
OUT = "/home/runner/work/project/project/coherent_synthesis_slides_v5.pptx"
prs.save(OUT)
print(f"✅  已生成：{OUT}")
