"""
generate_slides_v9.py
──────────────────────────────────────────────────────────────────────────
多节点高效相参合成总体研究框架图  ·  单页 PPT

布局（16:9 宽屏  13.33" × 7.50"）：
  ┌───────────────────────────────────────────────────────────────────┐
  │  标题栏（深蓝，与 v8 同款）                                       │
  ├───────────────────────────────────────────────────────────────────┤
  │  研究链路流程带：背景需求 ► 方法原理 ► 模型流程 ► 仿真实现        │
  ├──────────┬────────────────┬────────────────┬─────────────────────┤
  │  层次标签│  地基平台(绿) │  空中平台(蓝) │  空海平台(橙)       │
  │  ① 背景  │  ...          │  ...          │  ...                │
  │  ② 原理  │  ...          │  ...          │  ...                │
  │  ③ 流程  │  ...          │  ...          │  ...                │
  │  ④ 仿真  │  ...          │  ...          │  ...                │
  ├───────────────────────────────────────────────────────────────────┤
  │  高效相参合成核心指标（共性底栏）                                  │
  └───────────────────────────────────────────────────────────────────┘

风格：与 generate_slides_v8.py 完全一致（相同调色板、工具函数、排版规范）

运行：
    python generate_slides_v9.py
输出：
    coherent_synthesis_slides_v9.pptx
"""

import os
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# ━━━━━━━━━━━━━━━━━━━━━━ 调色板（与 v8 一致 + 平台扩展色）━━━━━━━━━━━━━
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
COL_GREEN   = RGBColor(0x0A, 0x5C, 0x36)

# 地基平台 —— 绿色系
GND_H   = RGBColor(0x1A, 0x5C, 0x38)   # 列头深绿
GND_R   = [RGBColor(0x2E, 0x7D, 0x52),
           RGBColor(0x33, 0x8A, 0x5E),
           RGBColor(0x3A, 0x99, 0x66),
           RGBColor(0x4B, 0xAE, 0x7E)]  # 各行
GND_BG  = RGBColor(0xEA, 0xF6, 0xEF)

# 空中平台 —— 蓝色系
AIR_H   = RGBColor(0x14, 0x35, 0xA0)
AIR_R   = [RGBColor(0x1F, 0x56, 0xC8),
           RGBColor(0x2D, 0x6E, 0xD8),
           RGBColor(0x3C, 0x82, 0xE0),
           RGBColor(0x50, 0x96, 0xE8)]
AIR_BG  = RGBColor(0xEB, 0xF0, 0xFF)

# 空海平台 —— 橙色系
SEA_H   = RGBColor(0x8B, 0x2E, 0x00)
SEA_R   = [RGBColor(0xB5, 0x42, 0x00),
           RGBColor(0xC8, 0x56, 0x10),
           RGBColor(0xDA, 0x6E, 0x20),
           RGBColor(0xE8, 0x86, 0x30)]
SEA_BG  = RGBColor(0xFF, 0xF0, 0xE8)

# 层次标签颜色（左列）
LBL_COLS = [
    RGBColor(0x2C, 0x5E, 0x8A),   # 背景需求
    RGBColor(0x1A, 0x5C, 0x38),   # 方法原理
    RGBColor(0x7A, 0x3B, 0x0E),   # 模型流程
    RGBColor(0x3B, 0x0E, 0x6E),   # 仿真实现
]
# 流程带（顶部4步）
FLOW_COLS = [
    RGBColor(0x2C, 0x5E, 0x8A),
    RGBColor(0x1A, 0x5C, 0x38),
    RGBColor(0x7A, 0x3B, 0x0E),
    RGBColor(0x3B, 0x0E, 0x6E),
]
# 底栏
BOT_COL = RGBColor(0x3B, 0x0E, 0x6E)

# ━━━━━━━━━━━━━━━━━━━━━━ 尺寸常量 ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
SW       = Inches(13.33)
SH       = Inches(7.50)
MARGIN   = Inches(0.30)
TITLE_H  = Inches(0.80)
CT       = TITLE_H + Inches(0.12)          # content top
FOOTER_Y = SH - Inches(0.44)
CB       = FOOTER_Y - Inches(0.04)
CW       = SW - 2 * MARGIN


# ━━━━━━━━━━━━━━━━━━━━━━ 工具函数（与 v8 同款）━━━━━━━━━━━━━━━━━━━━━━━

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


def tb_multi(slide, lines, l, t, w, h,
             sz=9, bold=False, color=DARK_TEXT,
             align=PP_ALIGN.LEFT, wrap=True):
    """Multi-paragraph textbox."""
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = wrap
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        r = p.add_run()
        r.text = ln
        r.font.size = Pt(sz)
        r.font.bold = bold
        r.font.color.rgb = color
    return box


def circle_badge(slide, cx, cy, r_inch, fill, text, sz=10, bold=True, tc=WHITE):
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


def title_bar(slide, title_main, title_sub, page, total=1):
    rect(slide, Inches(0), Inches(0), Inches(0.18), TITLE_H, COL_NAVY)
    rect(slide, Inches(0.18), Inches(0), SW - Inches(0.18), TITLE_H, TITLE_BAR)
    hline(slide, Inches(0), TITLE_H, SW, TITLE_BAR2, Pt(2.5))
    tb(slide, title_main,
       Inches(0.36), Inches(0.08), Inches(10.5), Inches(0.38),
       sz=19, bold=True, color=WHITE)
    tb(slide, title_sub,
       Inches(0.36), Inches(0.48), Inches(10.5), Inches(0.26),
       sz=10, color=RGBColor(0xB0, 0xCC, 0xF0))
    circle_badge(slide, SW - Inches(0.52), TITLE_H / 2, 0.26,
                 TITLE_BAR2, f"{page}/{total}", sz=10)


def footer(slide, text):
    rect(slide, Inches(0), FOOTER_Y, SW, SH - FOOTER_Y, FOOTER_BG)
    hline(slide, MARGIN, FOOTER_Y, SW - MARGIN, TITLE_BAR2, Pt(1.4))
    rect(slide, MARGIN, FOOTER_Y + Inches(0.09), Inches(0.05), Inches(0.25), TITLE_BAR2)
    tb(slide, "  " + text,
       MARGIN + Inches(0.10), FOOTER_Y + Inches(0.08),
       CW - Inches(0.14), Inches(0.32),
       sz=8.2, italic=True, color=COL_NAVY)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━ CONTENT DATA ━━━━━━━━━━━━━━━━━━━━━━━━━━━

# [row][col] => list of text lines
CELL_DATA = [
    # ── 行0：背景需求 ────────────────────────────────────────────────
    [
        ["■ 多固定站远程组网探测",
         "■ 时间/频率/相位同步约束",
         "■ 高增益低虚警检测需求",
         "■ 稳定基线·低机动性场景"],
        ["■ 机载编队协同作战",
         "■ 高速运动·基线时变",
         "■ 平台振动引入相位抖动",
         "■ 强杂波背景·实时处理"],
        ["■ 舰机异构协同探测",
         "■ 双/多基地构型变化",
         "■ 强海杂波·多径效应",
         "■ 异构量测配准需求"],
    ],
    # ── 行1：方法原理 ────────────────────────────────────────────────
    [
        ["◆ 相参合成增益理论模型",
         "◆ 最大比合并(MRC)加权",
         "◆ 增益退化律量化分析",
         "◆ 相位误差传播链路建模"],
        ["◆ 时变相位动态建模",
         "◆ EKF在线相位估计与预测",
         "◆ 运动补偿+多普勒配准",
         "◆ 自适应加权融合机制"],
        ["◆ 双基地信号等效转换",
         "◆ 异构量测统一域配准",
         "◆ 分层信号级融合策略",
         "◆ 海面散射相位特性建模"],
    ],
    # ── 行2：模型流程 ────────────────────────────────────────────────
    [
        ["① 多站回波同步采集",
         "② 脉冲压缩→距离对齐",
         "③ 相位校准→误差补偿",
         "④ MRC加权合并→检测"],
        ["① 平台运动状态估计",
         "② 多普勒域配准/补偿",
         "③ 时变相位预测更新",
         "④ 自适应融合→检测"],
        ["① 双基地RD图生成",
         "② 空间/时间坐标配准",
         "③ 统一域相参合并",
         "④ 联合CFAR检测"],
    ],
    # ── 行3：仿真实现 ────────────────────────────────────────────────
    [
        ["▶ 阵列相参增益蒙特卡洛验证",
         "▶ 相位误差灵敏度曲线",
         "▶ SINR提升量化(N节点律)",
         "▶ 同步精度对增益影响分析"],
        ["▶ 高动态相参合成增益验证",
         "▶ EKF相位估计均方误差",
         "▶ 与CRLB对比分析",
         "▶ 实时处理复杂度评估"],
        ["▶ 异构平台相参效能分析",
         "▶ 海杂波背景下相参增益",
         "▶ 配准误差对增益退化影响",
         "▶ 双基地探测概率曲线"],
    ],
]

ROW_LABELS = ["① 背景\n   需求", "② 方法\n   原理", "③ 模型\n   流程", "④ 仿真\n   实现"]
PLAT_TITLES = ["地基平台\nGround-based", "空中平台\nAirborne", "空海平台\nAir-Sea"]
PLAT_H_COLS = [GND_H, AIR_H, SEA_H]
PLAT_R_COLS = [GND_R, AIR_R, SEA_R]
PLAT_BG     = [GND_BG, AIR_BG, SEA_BG]


# ━━━━━━━━━━━━━━━━━━━━━━━━━━ MAIN ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

prs = Presentation()
prs.slide_width  = SW
prs.slide_height = SH
blank = prs.slide_layouts[6]

s1 = prs.slides.add_slide(blank)
set_bg(s1, BG)

# ── 1. 标题栏 ──────────────────────────────────────────────────────
title_bar(s1,
          "多节点高效相参合成  ·  总体研究框架",
          "Overall Research Framework for Efficient Multi-Node Coherent Synthesis  "
          "·  地基 / 空中 / 空海  三大平台场景",
          1)

# ── 2. 顶部研究链路流程带 ──────────────────────────────────────────
FLOW_TOP = CT
FLOW_H   = Inches(0.60)
flow_steps = [
    ("背景需求\n探测需求与约束分析",     FLOW_COLS[0]),
    ("方法原理\n高效相参合成机理",        FLOW_COLS[1]),
    ("模型流程\n端到端信号处理链路",      FLOW_COLS[2]),
    ("仿真实现\n性能量化与工程验证",      FLOW_COLS[3]),
]
ARR_W  = Inches(0.13)
n_f    = len(flow_steps)
usable = CW - ARR_W * (n_f - 1)
step_w = usable / n_f
BOX_TOP = FLOW_TOP + Inches(0.06)
BOX_H   = FLOW_H - Inches(0.10)

px = MARGIN
for i, (lbl, col) in enumerate(flow_steps):
    rect(s1, px, BOX_TOP, step_w, BOX_H, col,
         lc=WHITE, lw=Pt(0.5), rnd=True)
    tb(s1, lbl, px + Inches(0.06), BOX_TOP + Inches(0.02),
       step_w - Inches(0.12), BOX_H - Inches(0.04),
       sz=9, bold=True, align=PP_ALIGN.CENTER, color=WHITE)
    if i < n_f - 1:
        tb(s1, "►", px + step_w, BOX_TOP + BOX_H * 0.25,
           ARR_W, BOX_H * 0.50,
           sz=9.5, color=COL_NAVY, align=PP_ALIGN.CENTER)
    px += step_w + ARR_W

# ── 3. 计算主体网格布局 ────────────────────────────────────────────
BOT_STRIP_H = Inches(0.38)
BOT_Y       = FOOTER_Y - BOT_STRIP_H - Inches(0.04)

GRID_TOP   = FLOW_TOP + FLOW_H + Inches(0.08)
GRID_BOT   = BOT_Y - Inches(0.04)
GRID_H     = GRID_BOT - GRID_TOP

LBL_W      = Inches(0.70)       # 左侧层次标签列宽
COL_GAP    = Inches(0.10)       # 列间隙
PLAT_HDR_H = Inches(0.44)       # 平台列头高度
N_ROWS     = 4
ROW_GAP    = Inches(0.06)       # 行间隙
ROW_H      = (GRID_H - PLAT_HDR_H - ROW_GAP * (N_ROWS - 1)) / N_ROWS

DATA_COLS_W = CW - LBL_W - COL_GAP  # 3个数据列总宽（含2个列间隙）
COL_W = (DATA_COLS_W - 2 * COL_GAP) / 3

# 各列 x 起始
col_xs = [MARGIN + LBL_W + COL_GAP + i * (COL_W + COL_GAP) for i in range(3)]

# 各行 y 起始（从平台头下方开始）
plat_hdr_y = GRID_TOP
row_ys = []
y0 = GRID_TOP + PLAT_HDR_H
for i in range(N_ROWS):
    row_ys.append(y0)
    y0 += ROW_H + ROW_GAP

# ── 4. 平台列头 ────────────────────────────────────────────────────
for ci in range(3):
    cx0 = col_xs[ci]
    hcol = PLAT_H_COLS[ci]
    # 阴影
    rect(s1, cx0 + Inches(0.025), plat_hdr_y + Inches(0.025), COL_W, PLAT_HDR_H,
         RGBColor(0xB0, 0xB8, 0xC8))
    # 主背景
    rect(s1, cx0, plat_hdr_y, COL_W, PLAT_HDR_H, hcol, rnd=True)
    # 标题文字（两行）
    lines = PLAT_TITLES[ci].split('\n')
    tb(s1, lines[0],
       cx0, plat_hdr_y + Inches(0.04),
       COL_W, Inches(0.24),
       sz=13, bold=True, align=PP_ALIGN.CENTER, color=WHITE)
    tb(s1, lines[1],
       cx0, plat_hdr_y + Inches(0.26),
       COL_W, Inches(0.16),
       sz=8.5, align=PP_ALIGN.CENTER, color=RGBColor(0xCC, 0xE8, 0xFF))

# ── 5. 左侧层次标签列 ──────────────────────────────────────────────
for ri in range(N_ROWS):
    ry = row_ys[ri]
    lbl_col = LBL_COLS[ri]
    # 阴影
    rect(s1, MARGIN + Inches(0.025), ry + Inches(0.025), LBL_W, ROW_H,
         RGBColor(0xB0, 0xB8, 0xC8))
    # 主框
    rect(s1, MARGIN, ry, LBL_W, ROW_H, lbl_col, rnd=True)
    # 文字（两行）
    tb_multi(s1, ROW_LABELS[ri].split('\n'),
             MARGIN + Inches(0.04), ry + (ROW_H - Inches(0.46)) / 2,
             LBL_W - Inches(0.06), Inches(0.50),
             sz=9.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

# ── 6. 数据格 ──────────────────────────────────────────────────────
ACCENT_W = Inches(0.09)   # 左侧强调色竖条宽度
PAD_L    = Inches(0.14)
PAD_T    = Inches(0.06)
LINE_H   = Inches(0.165)  # 每行文字高度

for ri in range(N_ROWS):
    ry = row_ys[ri]
    for ci in range(3):
        cx0  = col_xs[ci]
        bg   = PLAT_BG[ci]
        acc  = PLAT_R_COLS[ci][ri]
        items = CELL_DATA[ri][ci]

        # 阴影
        rect(s1, cx0 + Inches(0.025), ry + Inches(0.025), COL_W, ROW_H,
             RGBColor(0xC0, 0xC8, 0xD8))
        # 格子背景
        rect(s1, cx0, ry, COL_W, ROW_H, bg,
             lc=acc, lw=Pt(0.8))
        # 左侧强调色竖条
        rect(s1, cx0, ry, ACCENT_W, ROW_H, acc, rnd=False)

        # 行数据文字
        n = len(items)
        text_h = ROW_H - 2 * PAD_T
        used_h = n * LINE_H
        y_start = ry + PAD_T + max(Inches(0), (text_h - used_h) / 2)
        for k, ln in enumerate(items):
            tb(s1, ln,
               cx0 + PAD_L, y_start + k * LINE_H,
               COL_W - PAD_L - Inches(0.06), LINE_H,
               sz=8.2, color=DARK_TEXT, align=PP_ALIGN.LEFT)

# ── 7. 底部核心指标共性栏 ─────────────────────────────────────────
rect(s1, Inches(0), BOT_Y, SW, BOT_STRIP_H, BOT_COL)

metrics = [
    "高效性核心指标：",
    "相参增益 G~N^2（理论上限）",
    " | 相位误差 sigma <= pi/8（3dB门限）",
    " | 复杂度 O(N*M)（线性可扩展）",
    " | 数据率 >= 相干时间倒数",
    " | 实时延迟 <= 雷达重复间隔",
]
metric_str = "  ".join(metrics)
tb(s1, metric_str,
   Inches(0.40), BOT_Y + Inches(0.06),
   SW - Inches(0.60), BOT_STRIP_H - Inches(0.10),
   sz=9.5, bold=False, color=WHITE, align=PP_ALIGN.LEFT)

# 调整页脚位置（覆盖底栏下方）
footer(s1,
       "研究路径：背景需求驱动 → 高效相参机理建模 → 三平台端到端处理流程 → "
       "仿真量化验证   ·   三大场景形成完整研究闭环")

# ── 8. 保存 ────────────────────────────────────────────────────────
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                   "coherent_synthesis_slides_v9.pptx")
prs.save(OUT)
print(f"Saved: {OUT}")
