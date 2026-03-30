"""
生成《空海平台多节点协同探测高效相参合成》优化版 PPT（v2）
依赖：python-pptx >= 0.6
运行：python generate_slides_v2.py
输出：coherent_synthesis_slides_v2.pptx

v2 排版优化要点：
  1. 流水线横条撑满全幅（原版右侧留白约 1.7"）
  2. 第2页四列宽度修正（原版溢出 ~0.07"）
  3. 右侧框图改为横向流程图，消除垂直堆叠的拥挤感
  4. 标题栏左侧加 4px 金色 Accent 条，增强专业感
  5. 各模块增加统一内边距，关键公式加高亮底色
  6. 页码指示器（右下角）
  7. 底部分隔线改为实色细线，替代斜体小字注释
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# ── 调色板 ────────────────────────────────────────────────
BG          = RGBColor(0x0D, 0x25, 0x4A)   # 深海蓝背景
PANEL       = RGBColor(0x12, 0x3A, 0x6E)   # 面板蓝
PANEL_DARK  = RGBColor(0x0A, 0x1E, 0x3C)   # 深面板（交替行）
GOLD        = RGBColor(0xF5, 0xC5, 0x18)   # 金色（标题/Accent）
ORANGE      = RGBColor(0xF0, 0x72, 0x28)   # 橙色强调
WHITE       = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT       = RGBColor(0xC8, 0xD8, 0xEA)   # 浅蓝灰（次要文字）
TEAL        = RGBColor(0x0C, 0x76, 0x76)   # 青色
PURPLE      = RGBColor(0x4A, 0x22, 0x7A)   # 紫色
BROWN       = RGBColor(0x7A, 0x3A, 0x10)   # 棕色
GREEN       = RGBColor(0x1A, 0x6B, 0x2E)   # 绿色
RED_DARK    = RGBColor(0x6B, 0x18, 0x18)   # 暗红
HIGHLIGHT   = RGBColor(0x00, 0x4A, 0x8F)   # 公式高亮蓝
SEPARATOR   = RGBColor(0x2A, 0x5A, 0x9A)   # 分隔线

# 各列主色（与原版保持一致）
COL_COLORS = [
    RGBColor(0x10, 0x5B, 0x9E),
    TEAL,
    PURPLE,
    BROWN,
]

# ── 幻灯片尺寸（16:9 宽屏）────────────────────────────────
SW = Inches(13.33)
SH = Inches(7.5)

# ── 通用几何常量 ──────────────────────────────────────────
MARGIN      = Inches(0.28)          # 左右页边距
TITLE_H     = Inches(0.82)          # 标题栏高度
CONTENT_TOP = TITLE_H + Inches(0.15)   # 内容区起始 Y
CONTENT_BOT = SH - Inches(0.45)    # 内容区底部 Y（留给注释行）
FOOTER_Y    = SH - Inches(0.4)     # 底部注释 Y
CW          = SW - 2 * MARGIN      # 内容区总宽 (12.77")

# ── 工具函数 ──────────────────────────────────────────────

def _fill_solid(shape, color):
    shape.fill.solid()
    shape.fill.fore_color.rgb = color


def set_bg(slide, color):
    bg = slide.background
    bg.fill.solid()
    bg.fill.fore_color.rgb = color


def add_rect(slide, l, t, w, h, fill, line_color=None, line_w=Pt(0.75), radius=False):
    """添加矩形或圆角矩形（radius=True 时圆角）"""
    shape_type = 5 if radius else 1   # 5=roundRect, 1=rect
    s = slide.shapes.add_shape(shape_type, l, t, w, h)
    _fill_solid(s, fill)
    if line_color:
        s.line.color.rgb = line_color
        s.line.width = line_w
    else:
        s.line.fill.background()
    return s


def add_line(slide, x1, y1, x2, y2=None, color=SEPARATOR, width=Pt(0.5)):
    """添加直线连接线。若 y2 为 None，则视为水平线（y2=y1）。"""
    if y2 is None:
        y2 = y1
    connector = slide.shapes.add_connector(1, x1, y1, x2, y2)
    connector.line.color.rgb = color
    connector.line.width = width
    return connector


def tb(slide, text, l, t, w, h,
       size=12, bold=False, color=WHITE,
       align=PP_ALIGN.LEFT, italic=False, wrap=True, spacing_after=0):
    """添加文本框，返回 shape"""
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = wrap
    para = tf.paragraphs[0]
    para.alignment = align
    if spacing_after:
        from pptx.util import Pt as _Pt
        para.space_after = _Pt(spacing_after)
    run = para.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    return box


def add_table(slide, rows, cols, l, t, w, h):
    return slide.shapes.add_table(rows, cols, l, t, w, h).table


def cell_fmt(cell, text, size=11, bold=False, bg=None,
             fg=WHITE, align=PP_ALIGN.CENTER):
    cell.text = text
    p = cell.text_frame.paragraphs[0]
    p.alignment = align
    r = p.runs[0] if p.runs else p.add_run()
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = fg
    if bg:
        cell.fill.solid()
        cell.fill.fore_color.rgb = bg


def title_bar(slide, title_text, page_num, total_pages=2):
    """统一绘制标题栏：深蓝底 + 左侧金色 accent 条 + 标题文字 + 右侧页码"""
    add_rect(slide, Inches(0), Inches(0), SW, TITLE_H, PANEL)
    # 左侧金色 accent 条（宽 8px ≈ 0.083"）
    add_rect(slide, Inches(0), Inches(0), Inches(0.1), TITLE_H, GOLD)
    # 标题文字
    tb(slide, title_text,
       Inches(0.22), Inches(0.1), Inches(11.5), Inches(0.62),
       size=21, bold=True, color=GOLD, align=PP_ALIGN.LEFT)
    # 右侧页码
    tb(slide, f"{page_num} / {total_pages}",
       Inches(12.4), Inches(0.24), Inches(0.8), Inches(0.38),
       size=11, color=LIGHT, align=PP_ALIGN.RIGHT)
    # 底部金色细线
    add_line(slide, Inches(0.12), TITLE_H, SW - Inches(0.12), color=GOLD, width=Pt(1.0))


def footer_note(slide, text):
    """底部分隔线 + 注释文字"""
    add_line(slide, MARGIN, FOOTER_Y - Inches(0.05),
             SW - MARGIN, color=SEPARATOR, width=Pt(0.5))
    tb(slide, text,
       MARGIN, FOOTER_Y, CW, Inches(0.38),
       size=9.5, italic=True, color=LIGHT)


# ══════════════════════════════════════════════════════════
#  构建演示文稿
# ══════════════════════════════════════════════════════════
prs = Presentation()
prs.slide_width  = SW
prs.slide_height = SH
blank = prs.slide_layouts[6]


# ╔══════════════════════════════════════════════════════╗
#  第 1 页：相参合成机理
# ╚══════════════════════════════════════════════════════╝
s1 = prs.slides.add_slide(blank)
set_bg(s1, BG)

title_bar(s1, "空海平台多节点协同探测  ·  高效相参合成 — 机理分析", 1)

# ── 布局参数 ──────────────────────────────────────────────
# 三列：左（机理卡片）| 中（公式）| 右（框图）
L_COL_W  = Inches(4.18)   # 左列宽
M_COL_W  = Inches(4.62)   # 中列宽
R_COL_W  = Inches(3.75)   # 右列宽（总=12.55，加两条间隙0.11*2=0.22，合计12.77=CW）
GAP      = Inches(0.11)

L_LEFT   = MARGIN
M_LEFT   = L_LEFT + L_COL_W + GAP
R_LEFT   = M_LEFT + M_COL_W + GAP
CT       = CONTENT_TOP  # 内容顶部
CB       = CONTENT_BOT  # 内容底部（不含注释行）
COL_H    = CB - CT      # 列内容高度

# 垂直分隔线
add_line(s1, M_LEFT - GAP/2, CT + Inches(0.1),
         M_LEFT - GAP/2, CB, SEPARATOR)
add_line(s1, R_LEFT - GAP/2, CT + Inches(0.1),
         R_LEFT - GAP/2, CB, SEPARATOR)

# ─────────────── 左列：四大耦合机理卡片 ───────────────────
mech_data = [
    ("① 时延同步",   COL_COLORS[0],
     "几何时延差 → 脉冲对齐补偿\n精度要求：Δτ < λ / (4c)\n方法：互相关峰值估计 + GPS授时"),
    ("② 相位同步",   TEAL,
     "载频相位差 → 复数旋转校正\n误差源：本振不一致、温漂\n方法：导频估计 / 回波自校准"),
    ("③ 频率同步",   PURPLE,
     "多普勒 + 振荡器频偏\n→ 相位线性漂移：ΔΦ = 2π·Δf·t\n方法：PLL / EKF 联合跟踪"),
    ("④ 幅度一致性", BROWN,
     "通道增益不均 → 旁瓣抬高\n→ MRC 权值补偿：wₖ ∝ SNRₖ\n方法：自适应归一化校准"),
]

card_h   = (COL_H - Inches(0.12) * 3) / 4   # 四卡等高，间隔 0.12"
for i, (title, col, body) in enumerate(mech_data):
    cy = CT + i * (card_h + Inches(0.12))
    # 卡片背景
    add_rect(s1, L_LEFT, cy, L_COL_W, card_h, col, radius=True)
    # 顶部色带（标题底色）
    add_rect(s1, L_LEFT, cy, L_COL_W, Inches(0.36), col, radius=True)
    # 标题文字
    tb(s1, title, L_LEFT + Inches(0.12), cy + Inches(0.04),
       L_COL_W - Inches(0.2), Inches(0.32),
       size=12, bold=True, color=GOLD)
    # 分隔线
    add_line(s1, L_LEFT + Inches(0.1), cy + Inches(0.38),
             L_LEFT + L_COL_W - Inches(0.1), cy + Inches(0.38),
             GOLD, Pt(0.5))
    # 正文（深色底板）
    add_rect(s1, L_LEFT, cy + Inches(0.38),
             L_COL_W, card_h - Inches(0.38),
             RGBColor(0x0A, 0x1E, 0x3C), radius=False)
    tb(s1, body,
       L_LEFT + Inches(0.12), cy + Inches(0.42),
       L_COL_W - Inches(0.2), card_h - Inches(0.48),
       size=10.5, color=WHITE, wrap=True)

# ─────────────── 中列：信号模型 + 增益退化公式 ────────────
MID_HALF = (COL_H - Inches(0.12)) / 2

# 上半：信号模型
add_rect(s1, M_LEFT, CT, M_COL_W, MID_HALF, PANEL, radius=True)
add_rect(s1, M_LEFT, CT, M_COL_W, Inches(0.36), PANEL, radius=True)
tb(s1, "信号模型",
   M_LEFT + Inches(0.14), CT + Inches(0.05),
   M_COL_W - Inches(0.2), Inches(0.3),
   size=13, bold=True, color=GOLD)
add_line(s1, M_LEFT + Inches(0.1), CT + Inches(0.38),
         M_LEFT + M_COL_W - Inches(0.1), CT + Inches(0.38), GOLD, Pt(0.5))

model_text = (
    "第 k 节点接收信号：\n"
    "  sₖ(t) = A · exp[j(2πfc·τₖ + φₖ)] + nₖ(t)\n\n"
    "相参合并输出：\n"
    "  y(t) = Σₖ  wₖ*· sₖ(t)\n\n"
    "最优相位对齐条件：\n"
    "  ∠wₖ = −(2πfc·τₖ + φₖ)"
)
tb(s1, model_text,
   M_LEFT + Inches(0.14), CT + Inches(0.44),
   M_COL_W - Inches(0.2), MID_HALF - Inches(0.52),
   size=11.5, color=WHITE, wrap=True)

# 下半：增益退化律
MY2 = CT + MID_HALF + Inches(0.12)
add_rect(s1, M_LEFT, MY2, M_COL_W, MID_HALF, PANEL, radius=True)
add_rect(s1, M_LEFT, MY2, M_COL_W, Inches(0.36), PANEL, radius=True)
tb(s1, "相参增益退化律（核心公式）",
   M_LEFT + Inches(0.14), MY2 + Inches(0.05),
   M_COL_W - Inches(0.2), Inches(0.3),
   size=13, bold=True, color=GOLD)
add_line(s1, M_LEFT + Inches(0.1), MY2 + Inches(0.38),
         M_LEFT + M_COL_W - Inches(0.1), MY2 + Inches(0.38), GOLD, Pt(0.5))

# 高亮公式底板
add_rect(s1, M_LEFT + Inches(0.1), MY2 + Inches(0.44),
         M_COL_W - Inches(0.2), Inches(0.48),
         HIGHLIGHT, line_color=GOLD, line_w=Pt(0.75))
tb(s1, "  G = N² · sinc²(σ_φ_total)",
   M_LEFT + Inches(0.14), MY2 + Inches(0.46),
   M_COL_W - Inches(0.28), Inches(0.44),
   size=13, bold=True, color=GOLD)

gain_body = (
    "联合相位误差方差：\n"
    "  σ²_φ = σ²_time·(2πfc)² + σ²_phase + σ²_freq·T²\n\n"
    "  σ_φ < π/8  →  G > 0.90·N²  （效能保持 90%）\n"
    "  σ_φ = π/4  →  G ≈ 0.64·N²  （退化 36%）"
)
tb(s1, gain_body,
   M_LEFT + Inches(0.14), MY2 + Inches(0.98),
   M_COL_W - Inches(0.2), MID_HALF - Inches(1.06),
   size=10.5, color=WHITE, wrap=True)

# ─────────────── 右列：增益对比 + 横向流程框图 ────────────
# 增益对比卡片（上）
GC_H = Inches(1.38)
add_rect(s1, R_LEFT, CT, R_COL_W, GC_H, PANEL, radius=True)
add_rect(s1, R_LEFT, CT, R_COL_W, Inches(0.36), PANEL, radius=True)
tb(s1, "增益对比",
   R_LEFT + Inches(0.12), CT + Inches(0.05),
   R_COL_W - Inches(0.2), Inches(0.3),
   size=13, bold=True, color=GOLD)
add_line(s1, R_LEFT + Inches(0.1), CT + Inches(0.38),
         R_LEFT + R_COL_W - Inches(0.1), CT + Inches(0.38), GOLD, Pt(0.5))
gain_cmp = (
    "相参合并：  SNR ∝ N²\n"
    "非相参合并：SNR ∝ N\n"
    "相参优势：额外 ×N 增益（+10·lgN dB）"
)
tb(s1, gain_cmp,
   R_LEFT + Inches(0.12), CT + Inches(0.44),
   R_COL_W - Inches(0.2), Inches(0.86),
   size=10.5, color=WHITE)

# 系统横向流程图（下部）
FD_TOP = CT + GC_H + Inches(0.14)
FD_H   = CB - FD_TOP
add_rect(s1, R_LEFT, FD_TOP, R_COL_W, FD_H, PANEL, radius=True)
add_rect(s1, R_LEFT, FD_TOP, R_COL_W, Inches(0.36), PANEL, radius=True)
tb(s1, "系统处理链路",
   R_LEFT + Inches(0.12), FD_TOP + Inches(0.05),
   R_COL_W - Inches(0.2), Inches(0.3),
   size=13, bold=True, color=GOLD)
add_line(s1, R_LEFT + Inches(0.1), FD_TOP + Inches(0.38),
         R_LEFT + R_COL_W - Inches(0.1), FD_TOP + Inches(0.38), GOLD, Pt(0.5))

# 框图节点（纵向排列，每节点一行）
flow_nodes = [
    ("节点 1", COL_COLORS[0]),
    ("节点 2", COL_COLORS[0]),
    ("节点 N", COL_COLORS[0]),
    ("时延对齐", TEAL),
    ("相位校正", PURPLE),
    ("MRC 合并 → N²", GREEN),
]
node_w = R_COL_W - Inches(0.28)
node_h = Inches(0.44)
n_gap  = (FD_H - Inches(0.44) - len(flow_nodes)*node_h) / (len(flow_nodes) - 1)
for j, (label, col) in enumerate(flow_nodes):
    ny = FD_TOP + Inches(0.44) + j * (node_h + n_gap)
    add_rect(s1, R_LEFT + Inches(0.14), ny, node_w, node_h, col,
             line_color=GOLD, line_w=Pt(0.75), radius=True)
    tb(s1, label,
       R_LEFT + Inches(0.14), ny, node_w, node_h,
       size=10, bold=(j == len(flow_nodes)-1), align=PP_ALIGN.CENTER)
    if j < len(flow_nodes) - 1:
        arr_y = ny + node_h + n_gap * 0.15
        tb(s1, "▼",
           R_LEFT + Inches(0.14) + node_w/2 - Inches(0.12),
           arr_y, Inches(0.24), n_gap * 0.7,
           size=9, color=GOLD, align=PP_ALIGN.CENTER)

# ── 底部注释 ──────────────────────────────────────────────
footer_note(s1,
    "★ 空海特殊挑战：载体运动导致基线动态变化，多普勒快变，"
    "四大误差耦合非线性退化，同步维持难度远高于地基系统。")


# ╔══════════════════════════════════════════════════════╗
#  第 2 页：相参合成方法设计
# ╚══════════════════════════════════════════════════════╝
s2 = prs.slides.add_slide(blank)
set_bg(s2, BG)

title_bar(s2, "空海平台多节点协同探测  ·  高效相参合成 — 方法设计", 2)

# ── 布局参数 ──────────────────────────────────────────────
PIPE_TOP = CONTENT_TOP
PIPE_H   = Inches(0.74)
COL_TOP  = PIPE_TOP + PIPE_H + Inches(0.15)
TBL_H    = Inches(1.70)
TBL_TOP  = FOOTER_Y - TBL_H - Inches(0.48)
# 四列区高度（流水线底 → 表格顶，减去标题行+间距）
FOUR_H   = TBL_TOP - COL_TOP - Inches(0.12)

# ── 流水线（全宽，撑满 CW）──────────────────────────────
pipe_steps = [
    ("① 原始回波\n输入",    COL_COLORS[0]),
    ("② 时延估计\n与对齐",  TEAL),
    ("③ 相位估计\n与校正",  PURPLE),
    ("④ 自适应\n权值计算",  BROWN),
    ("⑤ 相参合并\n输出",    GREEN),
    ("⑥ 目标检测\n处理",    RED_DARK),
]
n_steps  = len(pipe_steps)
arr_w    = Inches(0.20)   # 箭头宽
step_w   = (CW - (n_steps - 1) * arr_w) / n_steps   # 单步宽

for i, (label, col) in enumerate(pipe_steps):
    px = MARGIN + i * (step_w + arr_w)
    add_rect(s2, px, PIPE_TOP, step_w, PIPE_H, col,
             line_color=GOLD, line_w=Pt(0.75), radius=True)
    tb(s2, label, px, PIPE_TOP, step_w, PIPE_H,
       size=10.5, bold=True, align=PP_ALIGN.CENTER)
    if i < n_steps - 1:
        tb(s2, "▶",
           px + step_w, PIPE_TOP + PIPE_H * 0.18,
           arr_w, PIPE_H * 0.65,
           size=13, color=GOLD, align=PP_ALIGN.CENTER)

# ── 四列方法详解 ──────────────────────────────────────────
# 每列宽度：(CW - 3*gap) / 4，gap = 0.13"
COL_GAP  = Inches(0.13)
col_w    = (CW - 3 * COL_GAP) / 4

sections = [
    {
        "title": "① 时延对齐方法",
        "color": COL_COLORS[0],
        "items": [
            ("北斗/GPS 双频授时",         False),
            ("硬件同步精度 < 10 ns",      False),
            ("互相关峰值精估精度 ≈ 1/BW", False),
            ("INS/DVL 动态时延预测",      False),
            ("补偿载体机动时延抖动",      False),
            ("目标：Δτ < λ / (4c)",       True),
        ]
    },
    {
        "title": "② 相位校正方法",
        "color": TEAL,
        "items": [
            ("导频信号法（有导频）：",    False),
            ("  周期发送相位参考信号",    True),
            ("  LS 拟合载波相位漂移",     True),
            ("  实时反馈补偿",            True),
            ("回波自校准（无导频）：",    False),
            ("  利用多节点回波相位差",    True),
            ("  反解并补偿相位误差",      True),
            ("EKF / PLL 联合跟踪框架",   False),
        ]
    },
    {
        "title": "③ 自适应权值计算",
        "color": PURPLE,
        "items": [
            ("最大比合并（MRC）：",       False),
            ("  wₖ = SNRₖ / Σ SNRᵢ",    True),
            ("稳健加权：",                False),
            ("  相位误差大节点自动降权",  True),
            ("  防止相参合并被污染",      True),
            ("分层相参策略：",            False),
            ("  同步质量相近节点组内相参",True),
            ("  组间非相参叠加",          True),
        ]
    },
    {
        "title": "④ 空海特殊设计",
        "color": BROWN,
        "items": [
            ("海面多径处理：",            False),
            ("  分离直射/镜面反射分量",   True),
            ("  分量分离后再相参合并",    True),
            ("载体姿态补偿：",            False),
            ("  IMU 实时修正横摇/纵摇",   True),
            ("  Δφ=(2π/λ)·d·sin(Δθ)",   True),
            ("低截获设计：",              False),
            ("  相参合并在数字域完成",    True),
        ]
    },
]

TITLE_BAND_H = Inches(0.38)

for ci, sec in enumerate(sections):
    cx = MARGIN + ci * (col_w + COL_GAP)

    # 列标题背景
    add_rect(s2, cx, COL_TOP, col_w, TITLE_BAND_H, sec["color"],
             line_color=GOLD, line_w=Pt(0.75), radius=True)
    tb(s2, sec["title"], cx, COL_TOP, col_w, TITLE_BAND_H,
       size=11.5, bold=True, color=GOLD, align=PP_ALIGN.CENTER)

    # 列内容背景
    body_top = COL_TOP + TITLE_BAND_H
    body_h   = FOUR_H - TITLE_BAND_H
    add_rect(s2, cx, body_top, col_w, body_h, PANEL_DARK)

    # 逐行写内容
    item_h  = Inches(0.315)
    y_off   = body_top + Inches(0.06)
    for text, is_sub in sec["items"]:
        prefix = "    " if is_sub else "• "
        tb(s2, prefix + text,
           cx + Inches(0.10), y_off,
           col_w - Inches(0.15), item_h,
           size=9.5,
           color=LIGHT if is_sub else WHITE)
        y_off += item_h

# ── 性能对比表标题 ────────────────────────────────────────
tbl_label_y = TBL_TOP - Inches(0.40)
tb(s2, "▶  关键性能指标对比",
   MARGIN, tbl_label_y, Inches(5), Inches(0.36),
   size=12, bold=True, color=GOLD)
add_line(s2, MARGIN, tbl_label_y + Inches(0.36),
         MARGIN + CW, tbl_label_y + Inches(0.36), GOLD, Pt(0.75))

# ── 性能对比表 ────────────────────────────────────────────
tbl = add_table(s2, 4, 4, MARGIN, TBL_TOP, CW, TBL_H)

headers   = ["方法类型", "同步精度要求", "算法延迟", "适用场景"]
tbl_data  = [
    ["导频相参合并",   "时延 < 10 ns",  "低（< 1 ms）",    "协同组网、慢机动平台"],
    ["回波自校准合并", "时延 < 50 ns",  "中（1～10 ms）",  "无导频、通信受限场景"],
    ["分层相参合并",   "时延 < 100 ns", "低",               "节点数多、同步质量异构"],
]

for ci, h in enumerate(headers):
    cell_fmt(tbl.cell(0, ci), h, size=11, bold=True,
             bg=PANEL, fg=GOLD, align=PP_ALIGN.CENTER)

row_bgs = [RGBColor(0x12,0x2A,0x4A), PANEL_DARK, RGBColor(0x12,0x2A,0x4A)]
for ri, row in enumerate(tbl_data):
    for ci, val in enumerate(row):
        cell_fmt(tbl.cell(ri+1, ci), val, size=10.5,
                 bg=row_bgs[ri], fg=WHITE,
                 align=PP_ALIGN.LEFT if ci in (0, 3) else PP_ALIGN.CENTER)

# ── 底部注释 ──────────────────────────────────────────────
footer_note(s2,
    "★ 核心权衡：空海高动态平台要求在「估计精度 — 计算实时性 — 通信开销」"
    "三者之间取得最优平衡，并在极端机动场景下保持鲁棒性。")


# ── 保存 ──────────────────────────────────────────────────
output = "/home/runner/work/project/project/coherent_synthesis_slides_v2.pptx"
prs.save(output)
print(f"✅  优化版幻灯片已生成：{output}")
