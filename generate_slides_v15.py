"""
generate_slides_v15.py
─────────────────────────────────────────────────────────────────────────
空海平台多节点协同探测 — 相参合成两页 PPT（v15）

重新设计风格：「左侧导航栏 + 右侧内容区」现代简洁风
  ● 左侧暗色导航栏（宽 2.6"）：
    - 幻灯片编号 · 中文标题 · 英文副标题
    - 关键指标徽章（N/SNR增益/精度）
  ● 右侧白色内容区（宽 10.57"）：
    - 顶部全宽信号处理流程图
    - 主内容：左列（数据转换）+ 右列（信号级融合）
    - 每列 3 个卡片，采用左侧彩色边线 + 白色卡体（无阴影）
  ● 颜色方案：深蓝 #0D1B35 + 橙 #E55C00 + 蓝 #1155BB

运行：
    python generate_slides_v15.py
输出：
    coherent_synthesis_slides_v15.pptx
"""

import os
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# ━━━━━━━━━━━━━━━━━━━━━━ 尺寸常量 ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
SW = Inches(13.33)
SH = Inches(7.50)

SIDE_W   = Inches(2.60)        # 左侧导航栏宽度
SIDE_X   = Inches(0)
CT_X     = SIDE_W              # 内容区起点 X
CT_W     = SW - SIDE_W         # 内容区宽度  = 10.73"
MARGIN   = Inches(0.22)        # 内容区内边距

FOOTER_H = Inches(0.46)
FOOTER_Y = SH - FOOTER_H

FLOW_TOP = Inches(0.14)        # 流程图相对内容区顶部的偏移
FLOW_H   = Inches(0.82)
MAIN_TOP_OFF = FLOW_TOP + FLOW_H + Inches(0.12)
MAIN_H   = FOOTER_Y - MAIN_TOP_OFF - Inches(0.00)

# ━━━━━━━━━━━━━━━━━━━━━━ 调色板 ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 导航栏
SIDE_BG    = RGBColor(0x0D, 0x1B, 0x35)   # 极深海军蓝
SIDE_BG2   = RGBColor(0x16, 0x2A, 0x50)   # 稍浅
SIDE_ACCENT = RGBColor(0x2E, 0x5F, 0xBE)  # 蓝色强调线

# 内容区
CONT_BG    = RGBColor(0xF5, 0xF8, 0xFF)   # 极浅蓝白
CARD_BG    = RGBColor(0xFF, 0xFF, 0xFF)   # 纯白卡片
HDR_LINE   = RGBColor(0xD0, 0xDC, 0xF0)   # 卡片间分隔线

# 文字
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
DARK_TEXT  = RGBColor(0x1A, 0x20, 0x35)
MID_TEXT   = RGBColor(0x3A, 0x45, 0x65)
LIGHT_TEXT = RGBColor(0xA8, 0xC0, 0xE8)   # 导航栏副文字
FORMULA_TEXT = RGBColor(0x08, 0x40, 0x90)  # 公式蓝

# 主题色 — 数据转换（橙色系）
DT_MAIN  = RGBColor(0xE5, 0x5C, 0x00)   # 深橙 · 大区标题
DT_BADGE = RGBColor(0xFF, 0x80, 0x20)   # 浅橙 · 徽章
DT_LIGHT = RGBColor(0xFF, 0xF0, 0xE5)   # 极浅橙 · 卡片头背景

DT_S1    = RGBColor(0xE5, 0x5C, 0x00)   # 子块①
DT_S2    = RGBColor(0xCC, 0x55, 0x00)
DT_S3    = RGBColor(0xB3, 0x4E, 0x00)

# 主题色 — 信号级融合（蓝色系）
RD_MAIN  = RGBColor(0x11, 0x55, 0xBB)   # 蓝 · 大区标题
RD_BADGE = RGBColor(0x45, 0x88, 0xEE)   # 浅蓝 · 徽章
RD_LIGHT = RGBColor(0xEA, 0xF1, 0xFF)   # 极浅蓝 · 卡片头背景

RD_S1    = RGBColor(0x11, 0x55, 0xBB)
RD_S2    = RGBColor(0x00, 0x82, 0x7E)   # 青色
RD_S3    = RGBColor(0x62, 0x25, 0xAA)   # 紫

# 方法设计页
PIPE_BLUE   = RGBColor(0x11, 0x55, 0xBB)
PIPE_TEAL   = RGBColor(0x00, 0x82, 0x7E)
PIPE_PURPLE = RGBColor(0x62, 0x25, 0xAA)
PIPE_RED    = RGBColor(0xAA, 0x15, 0x15)
PIPE_GREEN  = RGBColor(0x14, 0x78, 0x30)
PIPE_ORANGE = RGBColor(0xC8, 0x55, 0x0A)

FORMULA_BG = RGBColor(0xE8, 0xF2, 0xFF)
FOOTER_BG  = RGBColor(0x0D, 0x1B, 0x35)


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


def hline(slide, x1, y, x2, color=HDR_LINE, width=Pt(0.8)):
    c = slide.shapes.add_connector(1, x1, y, x2, y)
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


def tb2(slide, lines_list, l, t, w, h, wrap=True):
    """lines_list: [(text, sz, bold, color, align)]"""
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = wrap
    first = True
    for (txt, sz, bold, col, align) in lines_list:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align
        r = p.add_run()
        r.text = txt
        r.font.size = Pt(sz)
        r.font.bold = bold
        r.font.color.rgb = col
    return box


def circle_badge(slide, cx, cy, r_inch, fill, text,
                 sz=11, bold=True, tc=WHITE):
    d = Inches(r_inch * 2)
    s = slide.shapes.add_shape(
        9, cx - Inches(r_inch), cy - Inches(r_inch), d, d)
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


# ── 导航栏 ─────────────────────────────────────────────────────────
def sidebar(slide, title_zh, title_en, page, total=2,
            metrics=None, section_labels=None):
    """Left dark navigation panel."""
    # Background
    rect(slide, SIDE_X, Inches(0), SIDE_W, SH, SIDE_BG)
    # Inner area slightly lighter
    rect(slide, SIDE_X + Inches(0.06), Inches(0),
         SIDE_W - Inches(0.06), SH, SIDE_BG2)
    # Top accent line
    rect(slide, SIDE_X, Inches(0), SIDE_W, Inches(0.06), SIDE_ACCENT)

    # Slide number circle
    circle_badge(slide,
                 SIDE_X + SIDE_W / 2, Inches(0.50), 0.28,
                 SIDE_ACCENT, f"{page}", sz=14, bold=True)
    tb(slide, f"/ {total}",
       SIDE_X + SIDE_W / 2 + Inches(0.27), Inches(0.38),
       Inches(0.40), Inches(0.24),
       sz=9, color=LIGHT_TEXT)

    # Title
    tb(slide, title_zh,
       SIDE_X + Inches(0.16), Inches(0.88),
       SIDE_W - Inches(0.24), Inches(0.70),
       sz=13, bold=True, color=WHITE, wrap=True)

    hline(slide,
          SIDE_X + Inches(0.14), Inches(1.62),
          SIDE_X + SIDE_W - Inches(0.14),
          color=SIDE_ACCENT, width=Pt(1.0))

    # English subtitle
    tb(slide, title_en,
       SIDE_X + Inches(0.14), Inches(1.68),
       SIDE_W - Inches(0.22), Inches(0.80),
       sz=7.5, color=LIGHT_TEXT, italic=True, wrap=True)

    # Section labels
    if section_labels:
        sy = Inches(2.62)
        for (dot_color, label_text) in section_labels:
            rect(slide, SIDE_X + Inches(0.16), sy + Inches(0.06),
                 Inches(0.06), Inches(0.06), dot_color)
            tb(slide, label_text,
               SIDE_X + Inches(0.30), sy,
               SIDE_W - Inches(0.38), Inches(0.22),
               sz=8.5, color=WHITE)
            sy += Inches(0.28)

    # Metrics boxes
    if metrics:
        my = SH - Inches(0.22) - len(metrics) * Inches(0.56)
        for (label, value, mc) in metrics:
            rect(slide, SIDE_X + Inches(0.14), my,
                 SIDE_W - Inches(0.28), Inches(0.50),
                 SIDE_BG, lc=mc, lw=Pt(1.2))
            rect(slide, SIDE_X + Inches(0.14), my,
                 Inches(0.06), Inches(0.50), mc)
            tb(slide, label,
               SIDE_X + Inches(0.26), my + Inches(0.04),
               SIDE_W - Inches(0.42), Inches(0.18),
               sz=7.5, color=LIGHT_TEXT)
            tb(slide, value,
               SIDE_X + Inches(0.26), my + Inches(0.22),
               SIDE_W - Inches(0.42), Inches(0.22),
               sz=9.5, bold=True, color=WHITE)
            my += Inches(0.56)

    # Bottom accent
    rect(slide, SIDE_X, SH - Inches(0.06), SIDE_W, Inches(0.06), SIDE_ACCENT)


# ── 内容区底部 footer ──────────────────────────────────────────────
def footer_bar(slide, text):
    rect(slide, CT_X, FOOTER_Y, CT_W, FOOTER_H, FOOTER_BG)
    hline(slide, CT_X, FOOTER_Y, SW,
          color=SIDE_ACCENT, width=Pt(1.5))
    rect(slide, CT_X + MARGIN, FOOTER_Y + Inches(0.10),
         Inches(0.05), Inches(0.26), SIDE_ACCENT)
    tb(slide, "  " + text,
       CT_X + MARGIN + Inches(0.10), FOOTER_Y + Inches(0.08),
       CT_W - MARGIN - Inches(0.20), Inches(0.30),
       sz=8.5, italic=True, color=LIGHT_TEXT)


# ── 内容区顶部：流程图 ────────────────────────────────────────────
def draw_pipeline(slide, steps, top, avail_w, h,
                  arrow_w=Inches(0.13)):
    n = len(steps)
    usable = avail_w - arrow_w * (n - 1)
    sw = usable / n
    bh = h - Inches(0.16)
    bt = top + Inches(0.10)
    px = CT_X + MARGIN
    for i, (lbl, col) in enumerate(steps):
        rect(slide, px, bt, sw, bh, col,
             lc=WHITE, lw=Pt(0.5), rnd=True)
        tb(slide, lbl, px, bt, sw, bh,
           sz=8.5, bold=True,
           align=PP_ALIGN.CENTER, color=WHITE)
        if i < n - 1:
            # Arrow
            tb(slide, "▶",
               px + sw, bt + bh * 0.28,
               arrow_w, bh * 0.44,
               sz=9, color=RGBColor(0x60, 0x80, 0xB0),
               align=PP_ALIGN.CENTER)
        px += sw + arrow_w


# ── 大区段落标题（横条）────────────────────────────────────────────
def section_bar(slide, l, t, w, h, fill, icon, title, subtitle=None,
                title_sz=13.5):
    SUB_CLR = RGBColor(0xFF, 0xE8, 0xCC) if fill == DT_MAIN else RGBColor(0xC0, 0xD8, 0xFF)
    rect(slide, l, t, w, h, fill)
    # left accent stripe
    rect(slide, l, t, Inches(0.08), h, WHITE)
    # icon
    tb(slide, icon,
       l + Inches(0.14), t + Inches(0.05),
       Inches(0.30), h - Inches(0.08),
       sz=15, bold=True, color=WHITE)
    # title
    tb(slide, title,
       l + Inches(0.52), t + Inches(0.04),
       w - Inches(0.60), h * 0.55,
       sz=title_sz, bold=True, color=WHITE)
    if subtitle:
        tb(slide, subtitle,
           l + Inches(0.52), t + h * 0.52,
           w - Inches(0.60), h * 0.46,
           sz=8, color=SUB_CLR)


# ── 内容卡片（左侧彩色边线风格）────────────────────────────────────
def content_card(slide, l, t, w, h, accent_color, badge_num,
                 title, items, title_sz=10.8, body_sz=9.0):
    """
    Flat card: thin left border accent + white body, no shadow.
    items: [(text, is_formula)]
    """
    BORDER_W = Inches(0.05)
    HDR_H    = Inches(0.32)

    # Card body
    rect(slide, l, t, w, h, CARD_BG,
         lc=RGBColor(0xD8, 0xE4, 0xF4), lw=Pt(0.6))
    # Left accent border
    rect(slide, l, t, BORDER_W, h, accent_color)
    # Header band (very light tint)
    hdr_fill = DT_LIGHT if accent_color in (DT_S1, DT_S2, DT_S3) else RD_LIGHT
    rect(slide, l + BORDER_W, t, w - BORDER_W, HDR_H, hdr_fill)
    # Separator under header
    hline(slide, l + BORDER_W, t + HDR_H, l + w,
          color=RGBColor(0xD0, 0xDC, 0xF0), width=Pt(0.6))

    # Badge circle
    circle_badge(slide,
                 l + BORDER_W + Inches(0.22), t + HDR_H / 2, 0.14,
                 accent_color, badge_num, sz=8.5, bold=True, tc=WHITE)
    # Title
    tb(slide, title,
       l + BORDER_W + Inches(0.44), t + Inches(0.04),
       w - BORDER_W - Inches(0.52), HDR_H - Inches(0.06),
       sz=title_sz, bold=True, color=accent_color)

    # Body items
    yy = t + HDR_H + Inches(0.04)
    item_h = body_sz * 1.85 * 914.4 / 72
    for (txt, is_formula) in items:
        clr   = FORMULA_TEXT if is_formula else DARK_TEXT
        prefix = "   " if is_formula else "• "
        tb(slide, prefix + txt,
           l + BORDER_W + Inches(0.08), yy,
           w - BORDER_W - Inches(0.14), item_h,
           sz=body_sz, color=clr)
        yy += item_h + Inches(0.008)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━ BUILD PPT ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

prs = Presentation()
prs.slide_width  = SW
prs.slide_height = SH
blank = prs.slide_layouts[6]


# ╔══════════════════════════════════════════════════════════════════════╗
#  第 1 页  ·  机理分析
# ╔══════════════════════════════════════════════════════════════════════╝
s1 = prs.slides.add_slide(blank)
set_bg(s1, CONT_BG)

# ── 左导航栏 ─────────────────────────────────────────────────────────
sidebar(s1,
        "相参合成\n机理分析",
        "Coherent Synthesis Mechanism:\nRange-Doppler Signal-Level Fusion",
        page=1, total=2,
        section_labels=[
            (SIDE_ACCENT,  "信号处理链路"),
            (DT_MAIN,      "数据转换配准"),
            (RD_MAIN,      "信号级融合"),
        ],
        metrics=[
            ("节点链路数量", "N = 6 链路",        SIDE_ACCENT),
            ("SNR 增益",     "15.6 dB（N=6）",    DT_MAIN),
            ("时延精度",     "< 0.5 ns",           RD_MAIN),
        ])

# ── 内容区背景 ────────────────────────────────────────────────────────
rect(s1, CT_X, Inches(0), CT_W, SH, CONT_BG)

# ── 分隔线（导航栏 | 内容区）────────────────────────────────────────
hline(s1, CT_X, Inches(0), CT_X, color=SIDE_ACCENT, width=Pt(0))
# Left edge accent
rect(s1, CT_X, Inches(0), Inches(0.04), SH, SIDE_ACCENT)

# ── 流程图 ────────────────────────────────────────────────────────────
avail_flow_w = CT_W - 2 * MARGIN
flow_steps = [
    ("原始回波\n(N=6链路)",    RGBColor(0x20, 0x60, 0xC0)),
    ("脉冲压缩\nPulse Comp",   RGBColor(0x20, 0x60, 0xC0)),
    ("距离多普勒图\nR-D Map",  RGBColor(0x20, 0x60, 0xC0)),
    ("数据转换\nData Conv",    DT_MAIN),
    ("信号级融合\nR-D Fusion", RD_MAIN),
    ("CFAR检测\n目标输出",     RGBColor(0x00, 0x82, 0x7E)),
]
draw_pipeline(s1, flow_steps,
              top=Inches(0.14),
              avail_w=avail_flow_w, h=FLOW_H)

# Dashed indicator lines under "数据转换" and "信号级融合"
step_tot_w = avail_flow_w
n_f = len(flow_steps)
ARR_W_F = Inches(0.13)
each_sw = (step_tot_w - ARR_W_F * (n_f - 1)) / n_f

dt_cx  = CT_X + MARGIN + 3 * (each_sw + ARR_W_F) + each_sw / 2
fus_cx = CT_X + MARGIN + 4 * (each_sw + ARR_W_F) + each_sw / 2
flow_b = Inches(0.10) + Inches(0.66)
for cx in [dt_cx, fus_cx]:
    c = s1.shapes.add_connector(
        1, cx, flow_b, cx, flow_b + Inches(0.16))
    c.line.color.rgb = RGBColor(0x88, 0xAA, 0xCC)
    c.line.width = Pt(1.2)

# ── 主内容区 ─────────────────────────────────────────────────────────
MAIN_T = Inches(0.14) + FLOW_H + Inches(0.14)
MAIN_B = FOOTER_Y - Inches(0.06)
MAIN_H_AVAIL = MAIN_B - MAIN_T
GAP_MID = Inches(0.14)

# 左右各占约 48%/52%（右侧标题字数多，略宽）
L_W = (avail_flow_w - GAP_MID) * 0.48
R_W = avail_flow_w - L_W - GAP_MID
L_X = CT_X + MARGIN
R_X = L_X + L_W + GAP_MID

SEC_H = Inches(0.44)

# ─ 左区大标题 ─
section_bar(s1, L_X, MAIN_T, L_W, SEC_H,
            DT_MAIN, "⚙", "数据转换   Data Conversion",
            subtitle="双基地量测 → 统一时延 / 多普勒 / 坐标域")

# ─ 右区大标题 ─
section_bar(s1, R_X, MAIN_T, R_W, SEC_H,
            RD_MAIN, "⚡", "距离多普勒域信号级融合",
            subtitle="Range-Doppler Domain Signal-Level Coherent Fusion")

# ─ 子卡片区域 ─
CARD_TOP = MAIN_T + SEC_H + Inches(0.10)
CARD_BOT = MAIN_B
CARD_TOT_H = CARD_BOT - CARD_TOP
SUB_GAP  = Inches(0.09)
CARD_H   = (CARD_TOT_H - 2 * SUB_GAP) / 3

dt_cards = [
    {
        "accent": DT_S1, "badge": "①",
        "title": "双基地时延配准（等效斜距折叠）",
        "items": [
            ("等效双程距离：r_eq = r_tx + r_rx  (双基地路径和)", False),
            ("r_fold = mod(r_eq, R_unamb)  ← 距离折叠处理",    True),
            ("rbin = round(r_fold / r_res) + 1  ← 距离门索引", True),
            ("精度要求：Δτ < λ/(4c) ≈ 0.5 ns ≡ Δr < 7.5 cm",  False),
        ]
    },
    {
        "accent": DT_S2, "badge": "②",
        "title": "多普勒域统一（速度转换）",
        "items": [
            ("双基地多普勒：f_d,k = (v_T·cosα_T + v_R·cosα_R) / λ", False),
            ("速度转换因子：K_v = cos(β/2)，β 为双基地角",           False),
            ("等效速度：v_eq = f_d·λ / (2·cos(β/2))",               True),
            ("各链路统一到等效单基地速度轴，消除多普勒偏差",         False),
        ]
    },
    {
        "accent": DT_S3, "badge": "③",
        "title": "坐标系配准与平台运动补偿",
        "items": [
            ("各平台坐标 → 统一 ECEF/ENU 参考系",               False),
            ("Δφ_motion = 2π·fc·Δτ_motion  ← 运动相位误差",     True),
            ("INS/DVL 辅助：精度 0.1 m/s、0.01°",               False),
            ("北斗/GPS 双频授时：时间基准同步 < 10 ns",          False),
        ]
    },
]

rd_cards = [
    {
        "accent": RD_S1, "badge": "①",
        "title": "RD图生成与信号模型",
        "items": [
            ("各链路：X_k[r,d] = CFFT{ s_k(t)·e^{-j2πf_d·t} }",         False),
            ("sₖ(t) = A·exp[j(2π·fc·τₖ + φₖ)] + nₖ(t)  ← 复数基带信号", True),
            ("信号级融合：在复数域（检测前）完成，保留完整相位信息",     False),
            ("区别于数据级融合：信噪比增益 N² vs N",                       False),
        ]
    },
    {
        "accent": RD_S2, "badge": "②",
        "title": "相参合并公式（MRC加权）",
        "items": [
            ("X_coh[r,d] = Σ_k  w_k · X_k[r,d] · exp(j·φ_k)",            True),
            ("MRC权值：w_k = SNR_k · coh_k · ph_cons_k  ← 三因子积",       False),
            ("相干度：coh_k = |Σ s_rot| / Σ|s_rot|  ← 回波自校准",         True),
            ("相位斜率一致性：ph_cons = exp(−|Δfd|/Δfd_ref)",               True),
        ]
    },
    {
        "accent": RD_S3, "badge": "③",
        "title": "增益退化律与效能分析",
        "items": [
            ("理想相参增益：G_coh = 20·lg(N) = 15.6 dB（N=6）",            False),
            ("非相参基准：G_incoh = 10·lg(N) = 7.8 dB，净增益 +7.8 dB",    False),
            ("含误差退化：G_eff = G · exp(−σ_φ²)",                          True),
            ("σ_φ < π/8 → G_eff ≥ 90%·G；σ_φ = π/4 → G_eff ≈ 64%·G",     False),
        ]
    },
]

for ci in range(3):
    cy = CARD_TOP + ci * (CARD_H + SUB_GAP)
    content_card(s1, L_X, cy, L_W, CARD_H,
                 dt_cards[ci]["accent"],
                 dt_cards[ci]["badge"],
                 dt_cards[ci]["title"],
                 dt_cards[ci]["items"])
    content_card(s1, R_X, cy, R_W, CARD_H,
                 rd_cards[ci]["accent"],
                 rd_cards[ci]["badge"],
                 rd_cards[ci]["title"],
                 rd_cards[ci]["items"])

footer_bar(s1,
           "核心机理：各链路双基地量测经数据转换配准至统一时延/多普勒/坐标域后，"
           "在距离多普勒复数域实施信号级相参合并，实现 N² 量级 SNR 增益")


# ╔══════════════════════════════════════════════════════════════════════╗
#  第 2 页  ·  相参合成方法设计
# ╔══════════════════════════════════════════════════════════════════════╝
s2 = prs.slides.add_slide(blank)
set_bg(s2, CONT_BG)

# ── 导航栏 ──────────────────────────────────────────────────────────
sidebar(s2,
        "相参合成\n方法设计",
        "Coherent Synthesis Method Design:\nProcessing Pipeline & Synchronization",
        page=2, total=2,
        section_labels=[
            (SIDE_ACCENT,  "两轮检测流水线"),
            (PIPE_BLUE,    "时延对齐"),
            (PIPE_TEAL,    "相位校正"),
            (PIPE_PURPLE,  "MRC合并"),
            (PIPE_ORANGE,  "三MTI + SIC"),
        ],
        metrics=[
            ("两轮级联检测", "Round 1 + Round 2",  SIDE_ACCENT),
            ("CFAR 门限",    "6 / 12 dB",           DT_MAIN),
            ("MTI 通道",     "3 通道并行",           RD_MAIN),
        ])

rect(s2, CT_X, Inches(0), CT_W, SH, CONT_BG)
rect(s2, CT_X, Inches(0), Inches(0.04), SH, SIDE_ACCENT)

# ── 两轮流水线 ────────────────────────────────────────────────────────
PIPE_TOP2 = Inches(0.14)
PIPE_H2   = Inches(0.80)

avail_flow_w2 = CT_W - 2 * MARGIN

# Round背景
rect(s2,
     CT_X + MARGIN, PIPE_TOP2,
     avail_flow_w2 * 0.50, PIPE_H2,
     RGBColor(0xDC, 0xEB, 0xF8),
     lc=RGBColor(0xB0, 0xC8, 0xE8), lw=Pt(0.6))
rect(s2,
     CT_X + MARGIN + avail_flow_w2 * 0.50, PIPE_TOP2,
     avail_flow_w2 * 0.50, PIPE_H2,
     RGBColor(0xD8, 0xEE, 0xE8),
     lc=RGBColor(0xA0, 0xCC, 0xBE), lw=Pt(0.6))
tb(s2, "Round 1 · 强目标检测（无MTI）",
   CT_X + MARGIN + Inches(0.10), PIPE_TOP2 + Inches(0.02),
   avail_flow_w2 * 0.48, Inches(0.18),
   sz=8, bold=True, color=PIPE_BLUE)
tb(s2, "Round 2 · 弱目标检测（三MTI并行）",
   CT_X + MARGIN + avail_flow_w2 * 0.50 + Inches(0.10), PIPE_TOP2 + Inches(0.02),
   avail_flow_w2 * 0.48, Inches(0.18),
   sz=8, bold=True, color=PIPE_TEAL)
c = s2.shapes.add_connector(
    1,
    CT_X + MARGIN + avail_flow_w2 * 0.50, PIPE_TOP2,
    CT_X + MARGIN + avail_flow_w2 * 0.50, PIPE_TOP2 + PIPE_H2)
c.line.color.rgb = RGBColor(0xF5, 0xC5, 0x18)
c.line.width = Pt(2.0)

pipe_steps2 = [
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
n_s     = len(pipe_steps2)
fracs   = [d[2] for d in pipe_steps2]
arr_tot = ARR_W2 * (n_s - 1)
usable2 = avail_flow_w2 - arr_tot
w_list  = [usable2 * f / sum(fracs) for f in fracs]

px2 = CT_X + MARGIN
BOX_TOP2 = PIPE_TOP2 + Inches(0.20)
BOX_H2   = PIPE_H2 - Inches(0.22)
for i, (lbl, col, _) in enumerate(pipe_steps2):
    sw = w_list[i]
    rect(s2, px2, BOX_TOP2, sw, BOX_H2, col,
         lc=WHITE, lw=Pt(0.5), rnd=True)
    tb(s2, lbl, px2, BOX_TOP2, sw, BOX_H2,
       sz=8, bold=True, align=PP_ALIGN.CENTER, color=WHITE)
    if i < n_s - 1:
        tb(s2, "▶", px2 + sw, BOX_TOP2 + BOX_H2 * 0.28,
           ARR_W2, BOX_H2 * 0.44,
           sz=9, color=RGBColor(0x50, 0x70, 0xA0),
           align=PP_ALIGN.CENTER)
    px2 += sw + ARR_W2

# ── 四列方法详解 ───────────────────────────────────────────────────
TBL_H    = Inches(1.50)
TBL_TOP  = FOOTER_Y - TBL_H - Inches(0.44)
M4_TOP   = PIPE_TOP2 + PIPE_H2 + Inches(0.16)
M4_H     = TBL_TOP - M4_TOP - Inches(0.12)
GAP4     = Inches(0.12)
COL4W    = (avail_flow_w2 - 3 * GAP4) / 4

method_sections = [
    {
        "num": "①", "color": PIPE_BLUE, "title": "时延对齐",
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
        "num": "②", "color": PIPE_TEAL, "title": "相位校正",
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
        "num": "③", "color": PIPE_PURPLE, "title": "MRC 合并",
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
        "num": "④", "color": PIPE_ORANGE, "title": "三MTI + SIC",
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

HEADER_H_CARD2 = Inches(0.38)
for ci, sec in enumerate(method_sections):
    cx = CT_X + MARGIN + ci * (COL4W + GAP4)
    # Card
    rect(s2, cx, M4_TOP, COL4W, M4_H, CARD_BG,
         lc=RGBColor(0xD0, 0xDC, 0xF0), lw=Pt(0.7))
    # Left accent border
    rect(s2, cx, M4_TOP, Inches(0.05), M4_H, sec["color"])
    # Header tint
    hdr_tint = RGBColor(0xEA, 0xF1, 0xFF)
    rect(s2, cx + Inches(0.05), M4_TOP,
         COL4W - Inches(0.05), HEADER_H_CARD2, hdr_tint)
    hline(s2, cx + Inches(0.05), M4_TOP + HEADER_H_CARD2, cx + COL4W,
          color=RGBColor(0xC0, 0xD0, 0xEC), width=Pt(0.6))
    # Badge
    circle_badge(s2,
                 cx + Inches(0.05) + Inches(0.22),
                 M4_TOP + HEADER_H_CARD2 / 2, 0.16,
                 sec["color"], sec["num"], sz=9, tc=WHITE)
    # Title
    tb(s2, sec["title"],
       cx + Inches(0.05) + Inches(0.46), M4_TOP + Inches(0.05),
       COL4W - Inches(0.60), HEADER_H_CARD2 - Inches(0.08),
       sz=12, bold=True, color=sec["color"])

    # Body items
    body_top = M4_TOP + HEADER_H_CARD2 + Inches(0.04)
    item_h = Inches(0.285)
    yy = body_top
    for (txt, is_code) in sec["items"]:
        clr    = FORMULA_TEXT if is_code else DARK_TEXT
        prefix = "   " if is_code else "• "
        tb(s2, prefix + txt,
           cx + Inches(0.12), yy,
           COL4W - Inches(0.20), item_h,
           sz=9.0, color=clr)
        yy += item_h

# ── 性能指标表 ─────────────────────────────────────────────────────
tbl_label_y = TBL_TOP - Inches(0.38)
rect(s2, CT_X + MARGIN, tbl_label_y, Inches(0.05), Inches(0.28), PIPE_BLUE)
tb(s2, "  关键性能指标 — 合并策略对比",
   CT_X + MARGIN + Inches(0.10), tbl_label_y,
   Inches(8), Inches(0.28),
   sz=11, bold=True, color=RGBColor(0x0A, 0x30, 0x70))
hline(s2, CT_X + MARGIN, tbl_label_y + Inches(0.30),
      CT_X + MARGIN + avail_flow_w2,
      color=PIPE_BLUE, width=Pt(1.5))

tbl = s2.shapes.add_table(
    5, 5,
    CT_X + MARGIN, TBL_TOP,
    avail_flow_w2, TBL_H).table


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


headers = ["合并策略", "同步精度要求", "算法延迟",
           "MATLAB 参数 / 函数", "适用场景"]
for ci, h in enumerate(headers):
    cf(tbl.cell(0, ci), h, sz=10, bold=True,
       bg=SIDE_BG, fg=WHITE)

tbl_data = [
    ["导频相参合并",   "时延 < 10 ns，Δφ < π/8",
     "< 1 ms",    "sys.fc=500M, cfar.min_snr_db=6",  "协同组网、慢机动"],
    ["回波自校准合并", "时延 < 50 ns",
     "1~10 ms",   "track_aided_coherent, ≥3 链路",    "无导频、通信受限"],
    ["三MTI 分层合并", "Δτ < 100 ns, coh ≥ 0.08",
     "低",         "mti_filter1/2/3, beta=0.65",       "节点多、同步异构"],
    ["SIC + 相参合并", "强目标 SNR > 12 dB",
     "两轮 ~2 ms","process_all_links_sic, sic_targets","强弱目标共存"],
]
alt_bgs = [FORMULA_BG, CARD_BG, FORMULA_BG, CARD_BG]
for ri, row in enumerate(tbl_data):
    for ci, val in enumerate(row):
        cf(tbl.cell(ri + 1, ci), val, sz=9.5,
           bg=alt_bgs[ri], fg=DARK_TEXT,
           al=PP_ALIGN.LEFT if ci in (0, 3, 4) else PP_ALIGN.CENTER)

footer_bar(s2,
           "核心权衡：空海高动态平台需在「估计精度 — 计算实时性 — 通信开销」三者间最优平衡；"
           "分层相参策略（SIC强目标 → 三MTI弱目标 → MRC合并）是算法设计核心")


# ━━━━━━━━━━━━━━━━━━━━━━ 保存 ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                   "coherent_synthesis_slides_v15.pptx")
prs.save(OUT)
print(f"✅  已生成：{OUT}")
