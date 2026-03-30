"""
generate_slides_v3.py
─────────────────────────────────────────────────────────────────────
依据 tx_air_v22_fixed_43_FINAL.m，生成两页独立 PPT：
  Page 1：空海平台多节点协同探测 — 高效相参合成机理
  Page 2：空海平台多节点协同探测 — 高效相参合成方法设计

版本亮点（v3）：
  • 内容直接来源于 MATLAB 代码（函数名、公式、参数数值均标注出处）
  • 第1页左右三栏：四大耦合误差卡片 | 信号模型+增益公式 | 系统参数+增益对比
  • 第2页上下结构：双轮处理流水线 | 四列方法详解 | 性能指标表
  • 统一主题色：深海蓝 + 金色标题 + 圆角卡片
运行：
    pip install python-pptx
    python generate_slides_v3.py
输出：
    coherent_synthesis_slides_v3.pptx
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# ── 调色板 ─────────────────────────────────────────────────────────
BG         = RGBColor(0x0B, 0x22, 0x45)   # 深海蓝背景
PANEL      = RGBColor(0x11, 0x38, 0x6A)   # 面板蓝
DARK_PANEL = RGBColor(0x08, 0x1C, 0x38)   # 深面板
GOLD       = RGBColor(0xF5, 0xC2, 0x18)   # 金色
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT      = RGBColor(0xC5, 0xD8, 0xEC)   # 浅蓝灰
ORANGE     = RGBColor(0xF0, 0x74, 0x28)   # 橙色
TEAL       = RGBColor(0x0B, 0x74, 0x74)
PURPLE     = RGBColor(0x49, 0x21, 0x78)
BROWN      = RGBColor(0x78, 0x38, 0x10)
GREEN      = RGBColor(0x19, 0x6A, 0x2C)
NAVY2      = RGBColor(0x0F, 0x59, 0x9C)   # 深蓝2
RED_DARK   = RGBColor(0x6A, 0x17, 0x17)
SEPARATOR  = RGBColor(0x28, 0x58, 0x98)
HIGHLIGHT  = RGBColor(0x00, 0x48, 0x8C)   # 公式高亮背景

# ── 尺寸常量（16:9） ───────────────────────────────────────────────
SW      = Inches(13.33)
SH      = Inches(7.5)
MARGIN  = Inches(0.28)
TITLE_H = Inches(0.82)
CT      = TITLE_H + Inches(0.14)   # content top
CB      = SH - Inches(0.44)        # content bottom (above footer)
FY      = SH - Inches(0.40)        # footer Y
CW      = SW - 2 * MARGIN          # content width ≈ 12.77"

# ── 工具函数 ───────────────────────────────────────────────────────

def set_bg(slide, color):
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = color


def rect(slide, l, t, w, h, fill, lc=None, lw=Pt(0.75), rnd=False):
    s = slide.shapes.add_shape(5 if rnd else 1, l, t, w, h)
    s.fill.solid(); s.fill.fore_color.rgb = fill
    if lc:
        s.line.color.rgb = lc; s.line.width = lw
    else:
        s.line.fill.background()
    return s


def line(slide, x1, y1, x2, y2=None, color=SEPARATOR, width=Pt(0.5)):
    if y2 is None: y2 = y1
    c = slide.shapes.add_connector(1, x1, y1, x2, y2)
    c.line.color.rgb = color; c.line.width = width


def tb(slide, text, l, t, w, h,
       sz=12, bold=False, color=WHITE,
       align=PP_ALIGN.LEFT, italic=False, wrap=True):
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame; tf.word_wrap = wrap
    p = tf.paragraphs[0]; p.alignment = align
    r = p.add_run(); r.text = text
    r.font.size = Pt(sz); r.font.bold = bold
    r.font.italic = italic; r.font.color.rgb = color
    return box


def title_bar(slide, title, page, total=2):
    rect(slide, Inches(0), Inches(0), SW, TITLE_H, PANEL)
    rect(slide, Inches(0), Inches(0), Inches(0.10), TITLE_H, GOLD)
    tb(slide, title,
       Inches(0.22), Inches(0.10), Inches(11.6), Inches(0.62),
       sz=20, bold=True, color=GOLD)
    tb(slide, f"{page} / {total}",
       Inches(12.38), Inches(0.24), Inches(0.82), Inches(0.35),
       sz=11, color=LIGHT, align=PP_ALIGN.RIGHT)
    line(slide, Inches(0.12), TITLE_H, SW - Inches(0.12),
         color=GOLD, width=Pt(0.9))


def footer(slide, text):
    line(slide, MARGIN, FY - Inches(0.06), SW - MARGIN,
         color=SEPARATOR, width=Pt(0.5))
    tb(slide, text, MARGIN, FY, CW, Inches(0.38),
       sz=9, italic=True, color=LIGHT)


def param_row(slide, label, value, lx, ty, row_h, col_w, bg, lbl_w_frac=0.52):
    """在参数表中绘制一行：标签左 | 值右"""
    lbl_w = col_w * lbl_w_frac
    val_w = col_w - lbl_w
    rect(slide, lx, ty, col_w, row_h, bg)
    line(slide, lx + lbl_w, ty, lx + lbl_w, ty + row_h,
         color=SEPARATOR, width=Pt(0.4))
    tb(slide, label, lx + Inches(0.06), ty + Inches(0.02),
       lbl_w - Inches(0.08), row_h - Inches(0.04), sz=9.5, color=LIGHT)
    tb(slide, value, lx + lbl_w + Inches(0.06), ty + Inches(0.02),
       val_w - Inches(0.08), row_h - Inches(0.04),
       sz=9.5, bold=True, color=WHITE)


# ══════════════════════════════════════════════════════════════════════
#  构建演示文稿
# ══════════════════════════════════════════════════════════════════════
prs = Presentation()
prs.slide_width  = SW
prs.slide_height = SH
blank = prs.slide_layouts[6]


# ╔════════════════════════════════════════════════════════════╗
#  第 1 页：相参合成机理
#  布局：左（四大耦合误差）| 中（信号模型 + 增益公式）| 右（系统参数 + 增益对比）
# ╚════════════════════════════════════════════════════════════╝
s1 = prs.slides.add_slide(blank)
set_bg(s1, BG)
title_bar(s1, "空海平台多节点协同探测  ·  高效相参合成 — 机理分析", 1)

# ── 列宽分配 ──────────────────────────────────────────────────────
L1W = Inches(4.20)   # 左列
M1W = Inches(4.60)   # 中列
R1W = CW - L1W - M1W - Inches(0.22)   # 右列  ≈ 3.75"
GAP1 = Inches(0.11)

L1 = MARGIN
M1 = L1 + L1W + GAP1
R1 = M1 + M1W + GAP1
COL_H = CB - CT

# 垂直分隔线
line(s1, M1 - GAP1/2, CT + Inches(0.1), M1 - GAP1/2, CB, SEPARATOR)
line(s1, R1 - GAP1/2, CT + Inches(0.1), R1 - GAP1/2, CB, SEPARATOR)

# ─────── 左列：四大耦合误差卡片（对应 MATLAB 系统误差源建模）────────
mech = [
    ("① 时延同步误差",  NAVY2,
     "几何双程时延差 → 脉冲错位\n"
     "精度要求：Δτ < λ/(4c)\n"
     "代码：r_fold = mod(r_eq, R_unamb)\n"
     "       tau = 2·r_fold / c\n"
     "方法：互相关峰值 + 亚像素插值\n"
     "       北斗/GPS 双频授时辅助"),
    ("② 相位同步误差",  TEAL,
     "载频相位不一致 → 复数旋转偏差\n"
     "代码：cell_c = cellv·exp(j·2π·fc·τ)\n"
     "       coh_sum += cell_c\n"
     "方法：导频估计 + EKF/PLL 联合跟踪\n"
     "       回波自校准（无导频场景）"),
    ("③ 频率同步误差",  PURPLE,
     "多普勒 + 振荡器频偏\n"
     "→ 相位线性漂移：ΔΦ = 2π·Δf·t\n"
     "代码：fd_phase = median(diff(unwrap(angle(s))))\n"
     "       ph_cons = exp(−|Δfd|/Δfd_ref)\n"
     "方法：相位斜率一致性检验 + PLL 锁定"),
    ("④ 幅度一致性",    BROWN,
     "通道增益不均 → 旁瓣抬高\n"
     "SNR 权值：score ∝ SNRₖ·coh·ph_cons\n"
     "代码：score = snr .* (0.3+0.7·coh)\n"
     "              .* (0.4+0.6·ph_cons)\n"
     "方法：MRC 自适应权值 + 自动降权"),
]

card_h = (COL_H - Inches(0.1) * 3) / 4
for i, (ttl, col, body) in enumerate(mech):
    cy = CT + i * (card_h + Inches(0.1))
    # 卡片外框
    rect(s1, L1, cy, L1W, card_h, col, lc=None, rnd=True)
    # 标题带
    rect(s1, L1, cy, L1W, Inches(0.34), col, rnd=True)
    tb(s1, ttl, L1 + Inches(0.12), cy + Inches(0.04),
       L1W - Inches(0.16), Inches(0.28), sz=11.5, bold=True, color=GOLD)
    line(s1, L1 + Inches(0.1), cy + Inches(0.36),
         L1 + L1W - Inches(0.1), cy + Inches(0.36), GOLD, Pt(0.45))
    # 内容深色底板
    rect(s1, L1, cy + Inches(0.37), L1W, card_h - Inches(0.37),
         DARK_PANEL)
    tb(s1, body,
       L1 + Inches(0.12), cy + Inches(0.40),
       L1W - Inches(0.18), card_h - Inches(0.46),
       sz=9.5, color=WHITE, wrap=True)

# ─────── 中列：信号模型 & 相参增益退化律 ────────────────────────────
MH = (COL_H - Inches(0.12)) / 2  # 上下各半

# ── 上半：信号模型
rect(s1, M1, CT, M1W, MH, PANEL, rnd=True)
rect(s1, M1, CT, M1W, Inches(0.34), PANEL, rnd=True)
tb(s1, "信号模型（基于 OFDM 双基地体制）",
   M1 + Inches(0.14), CT + Inches(0.04),
   M1W - Inches(0.2), Inches(0.28),
   sz=12, bold=True, color=GOLD)
line(s1, M1 + Inches(0.1), CT + Inches(0.36),
     M1 + M1W - Inches(0.1), CT + Inches(0.36), GOLD, Pt(0.45))

# 系统参数：fc=500 MHz, B=2 MHz, Nc=64, N=6
model_txt = (
    "第 k 节点（共 N=6 链路）接收信号：\n"
    "  sₖ(t) = A · exp[j(2π·fc·τₖ + φₖ)] + nₖ(t)\n\n"
    "  fc = 500 MHz，B = 2 MHz，PRI = 140 μs，Nc = 64\n\n"
    "相参合并输出：\n"
    "  coh_sum = Σₖ  cellₖ · exp(j·2π·fc·τₖ)\n"
    "  y = coh_sum  →  SNR ∝ N²\n\n"
    "最优对齐条件：∠wₖ = −(2π·fc·τₖ + φₖ)"
)
tb(s1, model_txt,
   M1 + Inches(0.14), CT + Inches(0.40),
   M1W - Inches(0.22), MH - Inches(0.48),
   sz=10.5, color=WHITE, wrap=True)

# ── 下半：相参增益退化律
MY2 = CT + MH + Inches(0.12)
rect(s1, M1, MY2, M1W, MH, PANEL, rnd=True)
rect(s1, M1, MY2, M1W, Inches(0.34), PANEL, rnd=True)
tb(s1, "相参增益估计（代码 track_aided_coherent）",
   M1 + Inches(0.14), MY2 + Inches(0.04),
   M1W - Inches(0.2), Inches(0.28),
   sz=12, bold=True, color=GOLD)
line(s1, M1 + Inches(0.1), MY2 + Inches(0.36),
     M1 + M1W - Inches(0.1), MY2 + Inches(0.36), GOLD, Pt(0.45))

# 高亮公式
rect(s1, M1 + Inches(0.1), MY2 + Inches(0.42),
     M1W - Inches(0.2), Inches(0.46),
     HIGHLIGHT, lc=GOLD, lw=Pt(0.7))
tb(s1, "  G = 10·lg( |Σ cellₖ·e^{j2πfcτₖ}|² / Σ|cellₖ|² )",
   M1 + Inches(0.14), MY2 + Inches(0.44),
   M1W - Inches(0.24), Inches(0.42),
   sz=12, bold=True, color=GOLD)

gain_txt = (
    "联合相位误差方差：\n"
    "  σ²_φ = σ²_time·(2π·fc)² + σ²_phase + σ²_freq·T²\n\n"
    "性能边界（代码 efficiency_boundary_report）：\n"
    "  σ_φ < π/8  →  G > 0.90·N²  （效能 ≥ 90%）\n"
    "  σ_φ = π/4  →  G ≈ 0.64·N²  （退化 36%）\n"
    "  实测阈值：used ≥ 3 链路方计算增益"
)
tb(s1, gain_txt,
   M1 + Inches(0.14), MY2 + Inches(0.94),
   M1W - Inches(0.22), MH - Inches(1.02),
   sz=10.5, color=WHITE, wrap=True)

# ─────── 右列：系统参数表 + 增益对比 ────────────────────────────────
# 系统参数表（来源 MATLAB sys 结构体）
tbl_top = CT
tbl_title_h = Inches(0.34)
rect(s1, R1, tbl_top, R1W, tbl_title_h, PANEL, rnd=True)
tb(s1, "系统参数（MATLAB sys.*）",
   R1 + Inches(0.12), tbl_top + Inches(0.04),
   R1W - Inches(0.16), tbl_title_h - Inches(0.06),
   sz=11, bold=True, color=GOLD)
line(s1, R1 + Inches(0.08), tbl_top + tbl_title_h,
     R1 + R1W - Inches(0.08), tbl_top + tbl_title_h, GOLD, Pt(0.45))

params = [
    ("sys.fc",     "500 MHz"),
    ("sys.B",      "2 MHz"),
    ("sys.Nsc",    "1024 子载波"),
    ("sys.PRI",    "140 μs"),
    ("sys.Nc",     "64 脉冲"),
    ("sys.r_res",  "75 m"),
    ("N 节点",     "2 Tx + 3 Rx = 6 链路"),
    ("sys.lambda", "0.6 m（c/500MHz）"),
]
row_h  = Inches(0.34)
row_bg = [RGBColor(0x0C, 0x24, 0x48), DARK_PANEL]
for j, (lbl, val) in enumerate(params):
    ry = tbl_top + tbl_title_h + j * row_h
    param_row(s1, lbl, val, R1, ry, row_h, R1W,
              row_bg[j % 2], lbl_w_frac=0.54)

# 增益对比卡片（下部）
gc_top = tbl_top + tbl_title_h + len(params) * row_h + Inches(0.10)
gc_h   = CB - gc_top
if gc_h > Inches(0.4):
    rect(s1, R1, gc_top, R1W, gc_h, PANEL, rnd=True)
    rect(s1, R1, gc_top, R1W, Inches(0.34), PANEL, rnd=True)
    tb(s1, "增益对比",
       R1 + Inches(0.12), gc_top + Inches(0.04),
       R1W - Inches(0.16), Inches(0.28),
       sz=11, bold=True, color=GOLD)
    line(s1, R1 + Inches(0.08), gc_top + Inches(0.36),
         R1 + R1W - Inches(0.08), gc_top + Inches(0.36), GOLD, Pt(0.45))
    gc_txt = (
        "相参合并：  SNR ∝ N²\n"
        "非相参合并：SNR ∝ N\n"
        "优势倍数：  ×N = ×6\n"
        "             = +7.8 dB"
    )
    tb(s1, gc_txt,
       R1 + Inches(0.12), gc_top + Inches(0.40),
       R1W - Inches(0.18), gc_h - Inches(0.46),
       sz=10.5, color=WHITE, wrap=True)

footer(s1,
    "★ 核心挑战：空海平台载体运动导致基线动态变化，多普勒快变，四大误差耦合非线性退化，"
    "同步维持难度远高于地基系统（sys.fc=500 MHz 时 λ=0.6 m，相位误差要求 < 47 mrad）。")


# ╔════════════════════════════════════════════════════════════╗
#  第 2 页：相参合成方法设计
#  布局：流水线（全宽）| 四列方法详解 | 性能指标表
# ╚════════════════════════════════════════════════════════════╝
s2 = prs.slides.add_slide(blank)
set_bg(s2, BG)
title_bar(s2, "空海平台多节点协同探测  ·  高效相参合成 — 方法设计", 2)

# ── 双轮处理流水线（全宽，来自 MATLAB 主循环）───────────────────────
PIPE_TOP = CT
PIPE_H   = Inches(0.70)

# Round 1（左半）+ Round 2（右半）两段流水线
round_defs = [
    # (label, fill_color, width_fraction)
    ("Round 1\n强目标检测\n（无MTI/积累）",  NAVY2,   0.16),
    ("CFAR\n粗检测\nSNR > 12 dB",           TEAL,    0.14),
    ("GN 优化\n定位解算",                    PURPLE,  0.12),
    ("SIC\n强目标对消",                      RED_DARK,0.10),
    ("Round 2\n弱目标检测\n（三MTI并行）",   NAVY2,   0.16),
    ("三MTI\n融合 + 去重\nCFAR SNR>6 dB",   TEAL,    0.14),
    ("GN 优化\n+RANSAC\n定位解算",           PURPLE,  0.10),
    ("KF 跟踪\n+ 相参增益\n估计",            GREEN,   0.08),
]

# 计算实际宽度（箭头宽 Inches(0.18)）
ARR_W   = Inches(0.17)
n_steps = len(round_defs)
fracs   = [d[2] for d in round_defs]
# 流水线总箭头占比
arr_total = ARR_W * (n_steps - 1)
usable = CW - arr_total
step_widths = [usable * f / sum(fracs) for f in fracs]

px = MARGIN
for i, (lbl, col, _) in enumerate(round_defs):
    sw = step_widths[i]
    rect(s2, px, PIPE_TOP, sw, PIPE_H, col, lc=GOLD, lw=Pt(0.7), rnd=True)
    tb(s2, lbl, px, PIPE_TOP, sw, PIPE_H,
       sz=9.5, bold=True, align=PP_ALIGN.CENTER)
    if i < n_steps - 1:
        tb(s2, "▶",
           px + sw, PIPE_TOP + PIPE_H * 0.15,
           ARR_W, PIPE_H * 0.70,
           sz=12, color=GOLD, align=PP_ALIGN.CENTER)
    px += sw + ARR_W

# 流水线分组标注（Round1 / Round2）
seps_x = MARGIN + sum(step_widths[:4]) + ARR_W * 3.5
line(s2, seps_x, PIPE_TOP - Inches(0.04),
     seps_x, PIPE_TOP + PIPE_H + Inches(0.04), GOLD, Pt(1.0))

# ── 四列方法详解 ─────────────────────────────────────────────────────
COL2_TOP  = PIPE_TOP + PIPE_H + Inches(0.14)
TBL2_H    = Inches(1.68)
TBL2_TOP  = FY - TBL2_H - Inches(0.46)
FOUR_H    = TBL2_TOP - COL2_TOP - Inches(0.10)
COLGAP    = Inches(0.12)
COL4W     = (CW - 3 * COLGAP) / 4
TITLE_BND = Inches(0.37)

sections = [
    {
        "title": "① 时延对齐",
        "color": NAVY2,
        "items": [
            ("北斗/GPS 双频授时（< 10 ns）",      False),
            ("互相关峰值 + 亚像素插值",            False),
            ("  r_fold = mod(r_eq, R_unamb)",      True),
            ("  rbin = round(r_fold/r_res)+1",     True),
            ("INS/DVL 动态预测载体运动",            False),
            ("精度目标：Δτ < λ/(4c) ≈ 0.5 ns",   False),
        ]
    },
    {
        "title": "② 相位校正",
        "color": TEAL,
        "items": [
            ("相干度估计（cfar_detect_1d）：",      False),
            ("  s_rot = s·exp(−j2π·fd0·PRI·n)",   True),
            ("  coh = |Σs_rot|/Σ|s_rot|",          True),
            ("相位斜率一致性（抑制镜像）：",        False),
            ("  ph = unwrap(angle(s))",             True),
            ("  fd_phase = median(diff(ph))/PRI",  True),
            ("  ph_cons = exp(−|Δfd|/Δfd_ref)",   True),
            ("EKF/PLL 实时跟踪相位漂移",           False),
        ]
    },
    {
        "title": "③ 权值与合并",
        "color": PURPLE,
        "items": [
            ("综合评分（代码 cfar_detect_1d）：",   False),
            ("  score = SNR",                       True),
            ("    × (0.3+0.7·coh)",               True),
            ("    × (0.4+0.6·ph_cons)",           True),
            ("MRC 合并：coh_sum = Σ cell·e^{jφ}", False),
            ("  incoh_sum = Σ |cell|²",            True),
            ("增益：G = 10·lg(|coh|²/incoh)",     False),
            ("分层相参：质量异构时组内相参",        False),
        ]
    },
    {
        "title": "④ 三MTI + 空海增强",
        "color": BROWN,
        "items": [
            ("三通道并行 MTI：",                    False),
            ("  SLOW  → 1阶: [1,−1]",              True),
            ("  FAST  → 2阶: [1,−2,1]",            True),
            ("  VFAST → 3阶: [1,−3,3,−1]",        True),
            ("速度分簇：SpeedSplit=[230,400] m/s", False),
            ("多帧能量积累：beta=0.65（powBank）",  False),
            ("SIC 强目标对消 → 弱目标提取",        False),
            ("姿态补偿：IMU 横摇/纵摇修正",        False),
        ]
    },
]

for ci, sec in enumerate(sections):
    cx = MARGIN + ci * (COL4W + COLGAP)
    # 标题带
    rect(s2, cx, COL2_TOP, COL4W, TITLE_BND, sec["color"],
         lc=GOLD, lw=Pt(0.7), rnd=True)
    tb(s2, sec["title"], cx, COL2_TOP, COL4W, TITLE_BND,
       sz=11, bold=True, color=GOLD, align=PP_ALIGN.CENTER)
    # 内容区
    body_top = COL2_TOP + TITLE_BND
    body_h   = FOUR_H - TITLE_BND
    rect(s2, cx, body_top, COL4W, body_h, DARK_PANEL)
    item_h = Inches(0.305)
    yy = body_top + Inches(0.06)
    for txt, is_sub in sec["items"]:
        prefix = "   " if is_sub else "• "
        tb(s2, prefix + txt,
           cx + Inches(0.10), yy,
           COL4W - Inches(0.14), item_h,
           sz=9.0, color=LIGHT if is_sub else WHITE)
        yy += item_h

# ── 性能指标表 ───────────────────────────────────────────────────────
tbl_lbl_y = TBL2_TOP - Inches(0.38)
tb(s2, "▶  关键性能指标（代码 efficiency_boundary_report + cfar 参数）",
   MARGIN, tbl_lbl_y, Inches(8), Inches(0.34),
   sz=11, bold=True, color=GOLD)
line(s2, MARGIN, tbl_lbl_y + Inches(0.34), MARGIN + CW,
     color=GOLD, width=Pt(0.75))

tbl = s2.shapes.add_table(5, 5, MARGIN, TBL2_TOP, CW, TBL2_H).table

def cf(cell, text, sz=10, bold=False, bg=None, fg=WHITE, al=PP_ALIGN.CENTER):
    cell.text = text
    p = cell.text_frame.paragraphs[0]; p.alignment = al
    r = p.runs[0] if p.runs else p.add_run()
    r.font.size = Pt(sz); r.font.bold = bold; r.font.color.rgb = fg
    if bg:
        cell.fill.solid(); cell.fill.fore_color.rgb = bg

headers = ["指标项", "同步精度要求", "算法延迟", "MATLAB 参数/函数", "适用场景"]
for ci, h in enumerate(headers):
    cf(tbl.cell(0, ci), h, sz=10.5, bold=True, bg=PANEL, fg=GOLD)

tbl_data = [
    ["导频相参合并",   "时延 < 10 ns，相位误差 < π/8",
     "低（< 1 ms）",  "sys.fc=500M，Nc=64，cfar.min_snr_db=6",  "协同组网、慢机动"],
    ["回波自校准合并", "时延 < 50 ns",
     "中（1~10 ms）", "track_aided_coherent，used ≥ 3 链路",     "无导频、通信受限"],
    ["三MTI分层合并",  "时延 < 100 ns，coh ≥ 0.08",
     "低",            "mti_filter1/2/3，beta=0.65，SpeedSplit",   "节点多、同步异构"],
    ["SIC+相参合并",   "强目标 SNR > 12 dB",
     "两轮（约2ms）", "process_all_links_sic，sic_targets",       "强弱目标共存"],
]
row_bgs = [RGBColor(0x0F, 0x2A, 0x50), DARK_PANEL,
           RGBColor(0x0F, 0x2A, 0x50), DARK_PANEL]
for ri, row in enumerate(tbl_data):
    for ci, val in enumerate(row):
        cf(tbl.cell(ri+1, ci), val, sz=9.5,
           bg=row_bgs[ri], fg=WHITE,
           al=PP_ALIGN.LEFT if ci in (0, 3, 4) else PP_ALIGN.CENTER)

footer(s2,
    "★ 核心权衡：空海高动态平台需在「估计精度 — 计算实时性 — 通信开销」三者间最优平衡；"
    "分层相参策略（强目标SIC → 三MTI弱目标 → MRC合并）是代码设计核心。")


# ── 保存 ──────────────────────────────────────────────────────────────
OUT = "/home/runner/work/project/project/coherent_synthesis_slides_v3.pptx"
prs.save(OUT)
print(f"✅  已生成：{OUT}")
