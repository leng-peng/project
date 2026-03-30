"""
generate_slides_v4.py
─────────────────────────────────────────────────────────────────────────
空海平台多节点协同探测 — 相参合成两页 PPT（美化版 v4）

  Page 1：相参合成 · 机理分析
  Page 2：相参合成 · 方法设计

美化亮点（v4 相较 v3）：
  • 每卡片加左侧色彩渐变 accent-bar，视觉层次更分明
  • 公式用高亮 spotlight 矩形 + 金色描边突出
  • 页面顶部标题栏加右侧圆形页码徽章
  • 流水线改为渐变箭头 + Round1/Round2 分组背景色块
  • 四列方法块加彩色序号圆徽章
  • 性能指标表头加渐变色背景，行交替色更和谐
  • CRLB 理论下界 + 增益对比用独立 spotlight 展示
  • 页脚加彩色左侧图标

运行：
    python generate_slides_v4.py
输出：
    coherent_synthesis_slides_v4.pptx
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# ━━━━━━━━━━━━━━━━━━━━━━ 调色板 ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
BG          = RGBColor(0x07, 0x18, 0x32)   # 深海蓝背景
PANEL       = RGBColor(0x0D, 0x2D, 0x5C)   # 主面板蓝
PANEL_DARK  = RGBColor(0x07, 0x1C, 0x3D)   # 深面板
PANEL_MID   = RGBColor(0x0F, 0x35, 0x68)   # 中深面板
GOLD        = RGBColor(0xF5, 0xC5, 0x18)   # 金色
GOLD_DARK   = RGBColor(0xB8, 0x8C, 0x00)   # 深金
WHITE       = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT       = RGBColor(0xB8, 0xD0, 0xE8)   # 浅蓝灰文字
SEP         = RGBColor(0x23, 0x4E, 0x88)   # 分割线
HIGHLIGHT   = RGBColor(0x00, 0x3A, 0x74)   # 公式背景

# 卡片 accent 色系
ACCENT_BLUE   = RGBColor(0x1A, 0x6F, 0xC8)
ACCENT_TEAL   = RGBColor(0x0B, 0x8E, 0x8E)
ACCENT_PURPLE = RGBColor(0x7B, 0x3A, 0xC0)
ACCENT_ORANGE = RGBColor(0xE8, 0x6A, 0x10)
ACCENT_GREEN  = RGBColor(0x15, 0x8A, 0x3E)
ACCENT_RED    = RGBColor(0xB5, 0x18, 0x18)

# 流水线模块色
PIPE_R1     = RGBColor(0x0F, 0x4A, 0x96)
PIPE_SIC    = RGBColor(0x7A, 0x18, 0x18)
PIPE_MTI    = RGBColor(0x0B, 0x72, 0x6A)
PIPE_GN     = RGBColor(0x5C, 0x20, 0x90)
PIPE_KF     = RGBColor(0x14, 0x74, 0x38)

# ━━━━━━━━━━━━━━━━━━━━━━ 尺寸常量 ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
SW      = Inches(13.33)
SH      = Inches(7.5)
MARGIN  = Inches(0.30)
TITLE_H = Inches(0.80)
CT      = TITLE_H + Inches(0.15)    # content top
CB      = SH - Inches(0.50)         # content bottom (above footer)
FY      = SH - Inches(0.46)         # footer Y
CW      = SW - 2 * MARGIN           # ≈ 12.73"
COL_H   = CB - CT


# ━━━━━━━━━━━━━━━━━━━━━━ 工具函数 ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def set_bg(slide, color):
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = color


def rect(slide, l, t, w, h, fill, lc=None, lw=Pt(0.75), rnd=False):
    shape_id = 5 if rnd else 1
    s = slide.shapes.add_shape(shape_id, l, t, w, h)
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    if lc:
        s.line.color.rgb = lc
        s.line.width = lw
    else:
        s.line.fill.background()
    return s


def line(slide, x1, y1, x2, y2=None, color=SEP, width=Pt(0.5)):
    if y2 is None:
        y2 = y1
    c = slide.shapes.add_connector(1, x1, y1, x2, y2)
    c.line.color.rgb = color
    c.line.width = width


def tb(slide, text, l, t, w, h,
       sz=12, bold=False, color=WHITE,
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


def tb_multi(slide, lines_list, l, t, w, h,
             default_sz=10, default_color=WHITE, wrap=True):
    """lines_list: [(text, sz, bold, color, align), ...]"""
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


def circle_badge(slide, cx, cy, r_inch, fill, text, sz=11, bold=True):
    """Draw a filled circle with centered text."""
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
    r.font.color.rgb = GOLD


def title_bar(slide, title, page, total=2):
    """Gradient-like title bar with page badge."""
    # Main bar
    rect(slide, Inches(0), Inches(0), SW, TITLE_H, PANEL)
    # Left accent stripe (thick)
    rect(slide, Inches(0), Inches(0), Inches(0.14), TITLE_H, GOLD)
    # Subtle inner stripe
    rect(slide, Inches(0.14), Inches(0), Inches(0.06), TITLE_H,
         RGBColor(0x18, 0x45, 0x88))
    # Title text
    tb(slide, title,
       Inches(0.32), Inches(0.12), Inches(11.2), Inches(0.56),
       sz=21, bold=True, color=GOLD)
    # Bottom gold divider
    line(slide, Inches(0.14), TITLE_H, SW - Inches(0.14),
         color=GOLD, width=Pt(1.0))
    # Page badge (circle)
    badge_cx = SW - Inches(0.58)
    badge_cy = TITLE_H / 2
    circle_badge(slide, badge_cx, badge_cy, 0.30,
                 RGBColor(0x18, 0x52, 0xA0),
                 f"{page}/{total}", sz=12)


def footer(slide, text, icon="★"):
    line(slide, MARGIN, FY - Inches(0.08), SW - MARGIN,
         color=SEP, width=Pt(0.6))
    # Icon accent
    rect(slide, MARGIN - Inches(0.02), FY, Inches(0.04),
         Inches(0.34), GOLD)
    tb(slide, f"  {icon}  {text}",
       MARGIN + Inches(0.06), FY + Inches(0.02),
       CW - Inches(0.08), Inches(0.36),
       sz=8.8, italic=True, color=LIGHT)


def card(slide, l, t, w, h, accent_color, title_text,
         body_lines, title_sz=11.5, body_sz=9.5,
         accent_w=Inches(0.07)):
    """
    Styled card with left accent bar, colored title band, dark body.
    body_lines: [(text, is_code)]
    """
    # Outer panel
    rect(slide, l, t, w, h, PANEL_MID, lc=RGBColor(0x20, 0x48, 0x80), lw=Pt(0.5))
    # Left accent bar
    rect(slide, l, t, accent_w, h, accent_color)
    # Title band
    title_band_h = Inches(0.37)
    rect(slide, l + accent_w, t, w - accent_w, title_band_h, accent_color)
    tb(slide, title_text,
       l + accent_w + Inches(0.10), t + Inches(0.04),
       w - accent_w - Inches(0.14), title_band_h - Inches(0.06),
       sz=title_sz, bold=True, color=GOLD)
    # Thin gold separator
    line(slide, l + accent_w + Inches(0.06), t + title_band_h,
         l + w - Inches(0.06), t + title_band_h, GOLD, Pt(0.5))
    # Body
    body_top = t + title_band_h + Inches(0.04)
    body_h = h - title_band_h - Inches(0.04)
    rect(slide, l + accent_w, body_top, w - accent_w, body_h, PANEL_DARK)
    yy = body_top + Inches(0.06)
    line_h = Inches(0.27)
    for (txt, is_code) in body_lines:
        clr = RGBColor(0x7A, 0xCF, 0xFF) if is_code else WHITE
        prefix = "   " if is_code else "• "
        tb(slide, prefix + txt,
           l + accent_w + Inches(0.10), yy,
           w - accent_w - Inches(0.16), line_h,
           sz=body_sz, color=clr)
        yy += line_h


def spotlight_box(slide, l, t, w, h, text,
                  sz=12, text_color=GOLD, border_color=GOLD):
    """Formula spotlight with dark fill and gold border."""
    rect(slide, l, t, w, h, HIGHLIGHT, lc=border_color, lw=Pt(1.2))
    tb(slide, text, l + Inches(0.12), t + Inches(0.04),
       w - Inches(0.22), h - Inches(0.08),
       sz=sz, bold=True, color=text_color, align=PP_ALIGN.CENTER)


# ════════════════════════════════════════════════════════════════════
#  构建演示文稿
# ════════════════════════════════════════════════════════════════════
prs = Presentation()
prs.slide_width  = SW
prs.slide_height = SH
blank = prs.slide_layouts[6]


# ╔══════════════════════════════════════════════════════════════════╗
#  第 1 页  ·  相参合成机理分析
#  布局：
#    左列（宽4.10"）：4大耦合误差卡片（accent 色条）
#    中列（宽4.65"）：信号模型 spotlight + 增益退化律 + CRLB
#    右列（宽3.70"）：系统参数表 + 增益对比 spotlight
# ╚══════════════════════════════════════════════════════════════════╝
s1 = prs.slides.add_slide(blank)
set_bg(s1, BG)
title_bar(s1, "空海平台多节点协同探测  ·  高效相参合成 — 机理分析", 1)

# ── 列宽分配 ────────────────────────────────────────────────────────
L1W = Inches(4.12)
M1W = Inches(4.65)
GAP = Inches(0.12)
R1W = CW - L1W - M1W - 2 * GAP

L1  = MARGIN
M1  = L1 + L1W + GAP
R1  = M1 + M1W + GAP

# 细垂直分隔线
line(s1, M1 - GAP/2, CT + Inches(0.06), M1 - GAP/2, CB, SEP)
line(s1, R1 - GAP/2, CT + Inches(0.06), R1 - GAP/2, CB, SEP)

# ─── 左列：4大耦合误差卡片 ──────────────────────────────────────────
error_cards = [
    (
        "① 时延同步误差",
        ACCENT_BLUE,
        [
            ("几何双程时延差 → 脉冲错位", False),
            ("精度要求：Δτ < λ/(4c) ≈ 0.5 ns", False),
            ("r_fold = mod(r_eq, R_unamb)", True),
            ("tau = 2·r_fold / c", True),
            ("方法：互相关 + 北斗/GPS 双频授时", False),
        ]
    ),
    (
        "② 相位同步误差",
        ACCENT_TEAL,
        [
            ("载频相位不一致 → 复数旋转偏差", False),
            ("cell_c = cellv·exp(j2π·fc·τ)", True),
            ("coh_sum += cell_c", True),
            ("方法：导频估计 + EKF/PLL 联合跟踪", False),
            ("无导频：回波自校准", False),
        ]
    ),
    (
        "③ 频率同步误差",
        ACCENT_PURPLE,
        [
            ("多普勒 + 振荡器频偏 → 相位漂移", False),
            ("ΔΦ = 2π · Δf · t（线性漂移）", False),
            ("ph_cons = exp(−|Δfd|/Δfd_ref)", True),
            ("方法：相位斜率一致性检验", False),
            ("PLL 实时跟踪频偏", False),
        ]
    ),
    (
        "④ 幅度一致性",
        ACCENT_ORANGE,
        [
            ("通道增益不均 → 合并旁瓣抬高", False),
            ("score = SNR × (0.3+0.7·coh)", True),
            ("       × (0.4+0.6·ph_cons)", True),
            ("方法：MRC 自适应加权合并", False),
            ("质量低链路自动降权", False),
        ]
    ),
]

card_gap  = Inches(0.09)
card_h    = (COL_H - card_gap * 3) / 4

for i, (title, accent, body_lines) in enumerate(error_cards):
    cy = CT + i * (card_h + card_gap)
    card(s1, L1, cy, L1W, card_h, accent, title, body_lines,
         title_sz=11.0, body_sz=9.2)

# ─── 中列：信号模型 + 增益退化律 + CRLB ─────────────────────────────
# ── 上段：系统配置 + 信号模型
seg1_h = COL_H * 0.38
rect(s1, M1, CT, M1W, seg1_h, PANEL_MID,
     lc=RGBColor(0x1A, 0x4A, 0x80), lw=Pt(0.5))
rect(s1, M1, CT, Inches(0.07), seg1_h, ACCENT_BLUE)
rect(s1, M1 + Inches(0.07), CT, M1W - Inches(0.07), Inches(0.37), ACCENT_BLUE)
tb(s1, "信号模型（OFDM 双基地体制）",
   M1 + Inches(0.18), CT + Inches(0.04),
   M1W - Inches(0.22), Inches(0.30),
   sz=12, bold=True, color=GOLD)
line(s1, M1 + Inches(0.10), CT + Inches(0.38),
     M1 + M1W - Inches(0.10), CT + Inches(0.38), GOLD, Pt(0.5))

# System config line
tb(s1, "系统：2 Tx（海面 15~18 m/s）+ 3 Rx（空中 3500~3800 m）= 6 双基地链路",
   M1 + Inches(0.12), CT + Inches(0.41),
   M1W - Inches(0.18), Inches(0.22),
   sz=9.5, color=LIGHT)

spotlight_box(
    s1,
    M1 + Inches(0.10), CT + Inches(0.65),
    M1W - Inches(0.18), Inches(0.42),
    "sₖ(t) = A · exp[ j(2π·fc·τₖ + φₖ) ] + nₖ(t)",
    sz=11.5,
)
tb(s1,
   "coh_sum = Σₖ  cellₖ · exp(j·2π·fc·τₖ)\n"
   "SNR_coh ∝ N²     SNR_incoh ∝ N",
   M1 + Inches(0.12), CT + Inches(1.12),
   M1W - Inches(0.18), Inches(0.40),
   sz=10, color=WHITE)

# ── 中段：增益退化律
seg2_top = CT + seg1_h + Inches(0.10)
seg2_h   = COL_H * 0.34
rect(s1, M1, seg2_top, M1W, seg2_h, PANEL_MID,
     lc=RGBColor(0x1A, 0x4A, 0x80), lw=Pt(0.5))
rect(s1, M1, seg2_top, Inches(0.07), seg2_h, ACCENT_TEAL)
rect(s1, M1 + Inches(0.07), seg2_top,
     M1W - Inches(0.07), Inches(0.37), ACCENT_TEAL)
tb(s1, "相参增益退化律（track_aided_coherent）",
   M1 + Inches(0.18), seg2_top + Inches(0.04),
   M1W - Inches(0.22), Inches(0.30),
   sz=12, bold=True, color=GOLD)
line(s1, M1 + Inches(0.10), seg2_top + Inches(0.38),
     M1 + M1W - Inches(0.10), seg2_top + Inches(0.38), GOLD, Pt(0.5))

spotlight_box(
    s1,
    M1 + Inches(0.10), seg2_top + Inches(0.42),
    M1W - Inches(0.18), Inches(0.40),
    "G = 10·lg( |Σ cellₖ·e^{j2πfcτₖ}|² / Σ|cellₖ|² )",
    sz=11,
)
gain_txt = (
    "σ²_φ = σ²_time·(2π·fc)² + σ²_phase + σ²_freq·T²\n"
    "σ_φ < π/8 → G > 0.90·N²（效能 ≥ 90%）\n"
    "σ_φ = π/4 → G ≈ 0.64·N²（退化 36%）"
)
tb(s1, gain_txt,
   M1 + Inches(0.12), seg2_top + Inches(0.86),
   M1W - Inches(0.18), Inches(0.50),
   sz=9.8, color=WHITE)

# ── 下段：CRLB 理论下界
seg3_top = seg2_top + seg2_h + Inches(0.10)
seg3_h   = CB - seg3_top
rect(s1, M1, seg3_top, M1W, seg3_h, PANEL_MID,
     lc=RGBColor(0x1A, 0x4A, 0x80), lw=Pt(0.5))
rect(s1, M1, seg3_top, Inches(0.07), seg3_h, ACCENT_PURPLE)
rect(s1, M1 + Inches(0.07), seg3_top,
     M1W - Inches(0.07), Inches(0.37), ACCENT_PURPLE)
tb(s1, "CRLB 理论下界（Fisher 信息矩阵）",
   M1 + Inches(0.18), seg3_top + Inches(0.04),
   M1W - Inches(0.22), Inches(0.30),
   sz=12, bold=True, color=GOLD)
line(s1, M1 + Inches(0.10), seg3_top + Inches(0.38),
     M1 + M1W - Inches(0.10), seg3_top + Inches(0.38), GOLD, Pt(0.5))

crlb_txt = (
    "测距误差下界：  σ_r  = c / (2B√(2·SNR))\n"
    "测速误差下界：  σ_v  = λ / (2Tc√SNR)\n"
    "FIM 联合累加： J = Σₖ HₖᵀRₖ⁻¹Hₖ\n"
    "6链路联合定位精度可降至 ~100 m 量级"
)
tb(s1, crlb_txt,
   M1 + Inches(0.12), seg3_top + Inches(0.42),
   M1W - Inches(0.18), seg3_h - Inches(0.50),
   sz=10, color=WHITE)

# ─── 右列：系统参数表 + 增益对比 ────────────────────────────────────
# 标题
rect(s1, R1, CT, R1W, Inches(0.37), PANEL_MID,
     lc=RGBColor(0x1A, 0x4A, 0x80), lw=Pt(0.5))
rect(s1, R1, CT, Inches(0.07), Inches(0.37), GOLD)
tb(s1, "系统参数（MATLAB sys.*）",
   R1 + Inches(0.14), CT + Inches(0.04),
   R1W - Inches(0.18), Inches(0.30),
   sz=11, bold=True, color=GOLD)
line(s1, R1 + Inches(0.07), CT + Inches(0.38),
     R1 + R1W - Inches(0.07), CT + Inches(0.38), GOLD, Pt(0.5))

params = [
    ("sys.fc",    "500 MHz"),
    ("sys.B",     "2 MHz"),
    ("sys.Nsc",   "1024 子载波"),
    ("sys.PRI",   "140 μs"),
    ("sys.Nc",    "64 脉冲"),
    ("sys.r_res", "75 m"),
    ("N 链路",    "2 Tx + 3 Rx = 6"),
    ("λ",         "0.6 m"),
]
row_h   = Inches(0.315)
row_bgs = [PANEL_MID, PANEL_DARK]
lbl_w   = R1W * 0.52
val_w   = R1W - lbl_w

for j, (lbl, val) in enumerate(params):
    ry = CT + Inches(0.39) + j * row_h
    bg = row_bgs[j % 2]
    rect(s1, R1, ry, R1W, row_h, bg)
    line(s1, R1 + lbl_w, ry, R1 + lbl_w, ry + row_h, SEP, Pt(0.4))
    tb(s1, lbl, R1 + Inches(0.10), ry + Inches(0.03),
       lbl_w - Inches(0.12), row_h - Inches(0.06), sz=9.5, color=LIGHT)
    tb(s1, val, R1 + lbl_w + Inches(0.06), ry + Inches(0.03),
       val_w - Inches(0.08), row_h - Inches(0.06), sz=9.5, bold=True, color=WHITE)

# 增益对比 spotlight
gc_top = CT + Inches(0.39) + len(params) * row_h + Inches(0.12)
gc_h   = CB - gc_top
if gc_h > Inches(0.6):
    rect(s1, R1, gc_top, R1W, gc_h, PANEL_MID,
         lc=GOLD, lw=Pt(0.7))
    rect(s1, R1, gc_top, Inches(0.07), gc_h, ACCENT_GREEN)
    rect(s1, R1 + Inches(0.07), gc_top,
         R1W - Inches(0.07), Inches(0.37), ACCENT_GREEN)
    tb(s1, "增益对比",
       R1 + Inches(0.18), gc_top + Inches(0.04),
       R1W - Inches(0.22), Inches(0.30),
       sz=11, bold=True, color=GOLD)
    line(s1, R1 + Inches(0.10), gc_top + Inches(0.38),
         R1 + R1W - Inches(0.10), gc_top + Inches(0.38), GOLD, Pt(0.5))
    spotlight_box(
        s1,
        R1 + Inches(0.10), gc_top + Inches(0.42),
        R1W - Inches(0.18), Inches(0.38),
        "相参 SNR ∝ N²  ÷  非相参 SNR ∝ N\n= ×N = ×6 = +7.8 dB",
        sz=10,
    )

footer(s1,
       "核心挑战：空海平台高动态下四大误差耦合非线性退化，"
       "同步要求 Δτ < 0.5 ns、Δφ < π/8（47 mrad），"
       "远超地基组网系统。",
       icon="▶")


# ╔══════════════════════════════════════════════════════════════════╗
#  第 2 页  ·  相参合成方法设计
#  布局：
#    顶部（全宽 0.72"）：双轮处理流水线
#    中部（四列等宽）  ：方法详解（带圆形序号徽章）
#    底部（全宽 1.62"）：关键性能指标表
# ╚══════════════════════════════════════════════════════════════════╝
s2 = prs.slides.add_slide(blank)
set_bg(s2, BG)
title_bar(s2, "空海平台多节点协同探测  ·  高效相参合成 — 方法设计", 2)

# ── 流水线区域 ──────────────────────────────────────────────────────
PIPE_TOP = CT
PIPE_H   = Inches(0.72)
PIPE_PAD = Inches(0.08)   # spacing between step boxes

# 背景分组：Round1 浅底 / Round2 浅底
r1_bg_w = CW * 0.50
r2_bg_x = MARGIN + r1_bg_w + Inches(0.06)
r2_bg_w = CW - r1_bg_w - Inches(0.06)
rect(s2, MARGIN, PIPE_TOP - Inches(0.04), r1_bg_w, PIPE_H + Inches(0.08),
     RGBColor(0x09, 0x22, 0x4A))
rect(s2, r2_bg_x, PIPE_TOP - Inches(0.04), r2_bg_w, PIPE_H + Inches(0.08),
     RGBColor(0x0A, 0x28, 0x40))
tb(s2, "Round 1  ·  强目标检测（无MTI）",
   MARGIN + Inches(0.08), PIPE_TOP - Inches(0.02),
   r1_bg_w - Inches(0.12), Inches(0.22),
   sz=8.5, color=LIGHT)
tb(s2, "Round 2  ·  弱目标检测（三MTI并行）",
   r2_bg_x + Inches(0.08), PIPE_TOP - Inches(0.02),
   r2_bg_w - Inches(0.12), Inches(0.22),
   sz=8.5, color=LIGHT)
# Round1/2 divider
line(s2, MARGIN + r1_bg_w + Inches(0.02),
     PIPE_TOP - Inches(0.06),
     MARGIN + r1_bg_w + Inches(0.02),
     PIPE_TOP + PIPE_H + Inches(0.06),
     GOLD, Pt(1.2))

pipe_steps = [
    # (label, color, frac)
    ("总回波输入\nN=6链路",            PIPE_R1,  0.12),
    ("无MTI积累\nCFAR SNR>12dB",      PIPE_R1,  0.12),
    ("GN 定位\n强目标解算",            PIPE_GN,  0.11),
    ("SIC\n强目标对消",                PIPE_SIC, 0.09),
    ("三MTI并行\nSLOW/FAST/VFAST",    PIPE_MTI, 0.14),
    ("CFAR SNR>6dB\nNMS去重",          PIPE_MTI, 0.12),
    ("RANSAC+GN\n精化定位",            PIPE_GN,  0.12),
    ("卡尔曼跟踪\n相参增益估计",       PIPE_KF,  0.12),
    ("目标输出\n位置/速度/ID",         PIPE_KF,  0.06),
]

ARR_W    = Inches(0.15)
n_s      = len(pipe_steps)
fracs    = [d[2] for d in pipe_steps]
arr_tot  = ARR_W * (n_s - 1)
usable   = CW - arr_tot
w_list   = [usable * f / sum(fracs) for f in fracs]

px = MARGIN
for i, (lbl, col, _) in enumerate(pipe_steps):
    sw = w_list[i]
    rect(s2, px, PIPE_TOP + Inches(0.20), sw, PIPE_H - Inches(0.20),
         col, lc=GOLD, lw=Pt(0.6), rnd=True)
    tb(s2, lbl,
       px, PIPE_TOP + Inches(0.20),
       sw, PIPE_H - Inches(0.20),
       sz=8.5, bold=True, align=PP_ALIGN.CENTER, color=WHITE)
    if i < n_s - 1:
        tb(s2, "▶",
           px + sw, PIPE_TOP + PIPE_H * 0.38,
           ARR_W, PIPE_H * 0.42,
           sz=11, color=GOLD, align=PP_ALIGN.CENTER)
    px += sw + ARR_W

# ── 四列方法详解 ────────────────────────────────────────────────────
COL2_TOP = PIPE_TOP + PIPE_H + Inches(0.16)
TBL_H    = Inches(1.55)
TBL_TOP  = FY - TBL_H - Inches(0.44)
FOUR_H   = TBL_TOP - COL2_TOP - Inches(0.10)
COLGAP   = Inches(0.12)
COL4W    = (CW - 3 * COLGAP) / 4
BADGE_R  = 0.24   # badge radius in inches
TITLE_BH = Inches(0.40)

method_sections = [
    {
        "num": "①",
        "title": "时延对齐",
        "color": ACCENT_BLUE,
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
        "num": "②",
        "title": "相位校正",
        "color": ACCENT_TEAL,
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
        "num": "③",
        "title": "MRC 合并",
        "color": ACCENT_PURPLE,
        "items": [
            ("综合评分（cfar_detect_1d）",      False),
            ("score = SNR × (0.3+0.7·coh)",     True),
            ("       × (0.4+0.6·ph_cons)",      True),
            ("MRC 合并：Σ cell·e^{jφ}",        False),
            ("G = 10·lg(|coh|²/incoh)",         True),
            ("分层相参：质量异构时组内相参",    False),
        ]
    },
    {
        "num": "④",
        "title": "三MTI + SIC",
        "color": ACCENT_ORANGE,
        "items": [
            ("三通道并行 MTI：",                False),
            ("SLOW  → [1,−1]   <230 m/s",       True),
            ("FAST  → [1,−2,1]  230~400",        True),
            ("VFAST → [1,−3,3,−1] >400",        True),
            ("多帧积累：beta=0.65 powBank",      False),
            ("SIC 强目标对消 → 弱目标提取",     False),
        ]
    },
]

for ci, sec in enumerate(method_sections):
    cx = MARGIN + ci * (COL4W + COLGAP)

    # Card background
    rect(s2, cx, COL2_TOP, COL4W, FOUR_H, PANEL_MID,
         lc=RGBColor(0x1A, 0x48, 0x80), lw=Pt(0.5))
    # Left accent
    rect(s2, cx, COL2_TOP, Inches(0.07), FOUR_H, sec["color"])
    # Title band
    rect(s2, cx + Inches(0.07), COL2_TOP,
         COL4W - Inches(0.07), TITLE_BH, sec["color"])
    tb(s2, sec["title"],
       cx + Inches(0.54), COL2_TOP + Inches(0.06),
       COL4W - Inches(0.60), TITLE_BH - Inches(0.08),
       sz=12, bold=True, color=GOLD)
    # Circular badge with number
    circle_badge(s2,
                 cx + Inches(0.07) + Inches(0.23),
                 COL2_TOP + TITLE_BH / 2,
                 BADGE_R,
                 RGBColor(0x05, 0x14, 0x2E),
                 sec["num"], sz=11)
    # Thin gold sep
    line(s2, cx + Inches(0.10), COL2_TOP + TITLE_BH,
         cx + COL4W - Inches(0.06), COL2_TOP + TITLE_BH, GOLD, Pt(0.5))
    # Body
    body_top = COL2_TOP + TITLE_BH + Inches(0.02)
    body_h   = FOUR_H - TITLE_BH
    rect(s2, cx + Inches(0.07), body_top,
         COL4W - Inches(0.07), body_h, PANEL_DARK)
    yy = body_top + Inches(0.06)
    item_h = Inches(0.29)
    for (txt, is_code) in sec["items"]:
        clr = RGBColor(0x7A, 0xCF, 0xFF) if is_code else WHITE
        prefix = "   " if is_code else "• "
        tb(s2, prefix + txt,
           cx + Inches(0.14), yy,
           COL4W - Inches(0.20), item_h,
           sz=9.2, color=clr)
        yy += item_h

# ── 性能指标表 ──────────────────────────────────────────────────────
tbl_lbl_y = TBL_TOP - Inches(0.38)
# Label with accent bar
rect(s2, MARGIN, tbl_lbl_y, Inches(0.07), Inches(0.30), GOLD)
tb(s2, "  关键性能指标（efficiency_boundary_report + cfar 参数）",
   MARGIN + Inches(0.10), tbl_lbl_y,
   Inches(9), Inches(0.30),
   sz=11, bold=True, color=GOLD)
line(s2, MARGIN, tbl_lbl_y + Inches(0.32), MARGIN + CW,
     color=GOLD, width=Pt(0.8))

tbl = s2.shapes.add_table(5, 5, MARGIN, TBL_TOP, CW, TBL_H).table


def cf(cell, text, sz=10, bold=False, bg=None, fg=WHITE, al=PP_ALIGN.CENTER):
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


# Header row with styled background
hdr_bg = RGBColor(0x0F, 0x3E, 0x7A)
headers = ["合并策略", "同步精度要求", "算法延迟", "MATLAB 参数 / 函数", "适用场景"]
for ci, h in enumerate(headers):
    cf(tbl.cell(0, ci), h, sz=10.5, bold=True, bg=hdr_bg, fg=GOLD)

tbl_data = [
    ["导频相参合并",    "时延 < 10 ns，Δφ < π/8",
     "< 1 ms",     "sys.fc=500M，Nc=64，cfar.min_snr_db=6",  "协同组网、慢机动"],
    ["回波自校准合并",  "时延 < 50 ns",
     "1~10 ms",    "track_aided_coherent，used ≥ 3 链路",     "无导频、通信受限"],
    ["三MTI 分层合并",  "时延 < 100 ns，coh ≥ 0.08",
     "低",          "mti_filter1/2/3，beta=0.65，SpeedSplit",   "节点多、同步异构"],
    ["SIC + 相参合并",  "强目标 SNR > 12 dB",
     "两轮 ~2 ms", "process_all_links_sic，sic_targets",       "强弱目标共存"],
]
row_bgs = [RGBColor(0x0D, 0x2A, 0x54), PANEL_DARK,
           RGBColor(0x0D, 0x2A, 0x54), PANEL_DARK]
for ri, row in enumerate(tbl_data):
    for ci, val in enumerate(row):
        cf(tbl.cell(ri + 1, ci), val, sz=9.5,
           bg=row_bgs[ri], fg=WHITE,
           al=PP_ALIGN.LEFT if ci in (0, 3, 4) else PP_ALIGN.CENTER)

footer(s2,
       "核心权衡：空海高动态平台需在「估计精度 — 计算实时性 — 通信开销」三者间最优平衡；"
       "分层相参策略（SIC强目标 → 三MTI弱目标 → MRC合并）是算法设计核心。",
       icon="▶")


# ── 保存 ────────────────────────────────────────────────────────────
OUT = "/home/runner/work/project/project/coherent_synthesis_slides_v4.pptx"
prs.save(OUT)
print(f"✅  已生成：{OUT}")
