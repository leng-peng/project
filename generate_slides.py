"""
生成《空海平台多节点协同探测高效相参合成》两页PPT
依赖：python-pptx >= 0.6
运行：python generate_slides.py
输出：coherent_synthesis_slides.pptx
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
import copy

# ── 颜色常量 ──────────────────────────────────────────────
DARK_BLUE   = RGBColor(0x0D, 0x2B, 0x55)   # 深蓝背景
MID_BLUE    = RGBColor(0x1A, 0x4F, 0x8A)   # 次蓝（色块）
ACCENT_GOLD = RGBColor(0xF5, 0xB8, 0x00)   # 金色标题
ACCENT_ORG  = RGBColor(0xF0, 0x6E, 0x22)   # 橙色强调
WHITE       = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT_GREY  = RGBColor(0xCC, 0xD6, 0xE0)
TABLE_HEADER= RGBColor(0x1A, 0x4F, 0x8A)

# ── 工具函数 ──────────────────────────────────────────────

def set_bg(slide, color: RGBColor):
    """设置幻灯片纯色背景"""
    from pptx.oxml.ns import qn
    from lxml import etree
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_textbox(slide, text, left, top, width, height,
                font_size=14, bold=False, color=WHITE,
                align=PP_ALIGN.LEFT, italic=False, wrap=True):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = wrap
    para = tf.paragraphs[0]
    para.alignment = align
    run = para.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    return txBox


def add_rect(slide, left, top, width, height, fill_color, line_color=None, line_width=Pt(0)):
    from pptx.util import Pt as PtU
    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE_TYPE.RECTANGLE
        left, top, width, height
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if line_color:
        shape.line.color.rgb = line_color
        shape.line.width = line_width
    else:
        shape.line.fill.background()
    return shape


def add_table(slide, rows, cols, left, top, width, height):
    table_shape = slide.shapes.add_table(rows, cols, left, top, width, height)
    return table_shape.table


def cell_style(cell, text, font_size=11, bold=False, bg_color=None,
               font_color=WHITE, align=PP_ALIGN.CENTER):
    cell.text = text
    para = cell.text_frame.paragraphs[0]
    para.alignment = align
    run = para.runs[0] if para.runs else para.add_run()
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = font_color
    if bg_color:
        cell.fill.solid()
        cell.fill.fore_color.rgb = bg_color


# ── 幻灯片尺寸（16:9 宽屏）────────────────────────────────
SW = Inches(13.33)
SH = Inches(7.5)

prs = Presentation()
prs.slide_width  = SW
prs.slide_height = SH

blank_layout = prs.slide_layouts[6]   # 空白版式

# ══════════════════════════════════════════════════════════
#  第 1 页：相参合成机理
# ══════════════════════════════════════════════════════════
slide1 = prs.slides.add_slide(blank_layout)
set_bg(slide1, DARK_BLUE)

# ── 顶部标题栏 ──
add_rect(slide1, Inches(0), Inches(0), SW, Inches(0.85), MID_BLUE)
add_textbox(slide1,
            "空海平台多节点协同探测  |  高效相参合成机理",
            Inches(0.3), Inches(0.08), Inches(12), Inches(0.65),
            font_size=22, bold=True, color=ACCENT_GOLD, align=PP_ALIGN.LEFT)

# ── 左侧：四大耦合机理色块 ──
mechanisms = [
    ("① 时延同步",    "目标-节点几何时延差\n→ 脉冲对齐补偿\n→ 精度要求：Δτ < λ/(4c)"),
    ("② 相位同步",    "载频相位差\n→ 复数旋转校正\n→ 误差源：振荡器不一致"),
    ("③ 频率同步",    "多普勒偏移 + 振荡器偏差\n→ 相位随时间漂移\n→ ΔΦ = 2π·Δf·t"),
    ("④ 幅度一致性",  "通道增益差异\n→ 合并权值补偿\n→ 影响旁瓣抑制性能"),
]
colors_m = [RGBColor(0x10,0x5B,0x9E), RGBColor(0x0E,0x6E,0x6E),
            RGBColor(0x4A,0x23,0x7A), RGBColor(0x7A,0x3A,0x10)]

for i, (title, body) in enumerate(mechanisms):
    top = Inches(1.05) + i * Inches(1.55)
    add_rect(slide1, Inches(0.25), top, Inches(4.1), Inches(1.45), colors_m[i])
    add_textbox(slide1, title,
                Inches(0.35), top + Inches(0.08), Inches(3.9), Inches(0.35),
                font_size=13, bold=True, color=ACCENT_GOLD)
    add_textbox(slide1, body,
                Inches(0.35), top + Inches(0.40), Inches(3.9), Inches(0.98),
                font_size=11, color=WHITE)

# ── 中央：信号模型与增益公式 ──
add_rect(slide1, Inches(4.65), Inches(1.05), Inches(4.5), Inches(2.9), MID_BLUE)
add_textbox(slide1, "信号模型",
            Inches(4.8), Inches(1.1), Inches(4.2), Inches(0.38),
            font_size=14, bold=True, color=ACCENT_GOLD)
model_text = (
    "第 k 节点接收信号：\n"
    "  sₖ(t) = A · exp[j(2πfcτₖ + φₖ)] + nₖ(t)\n\n"
    "相参合并输出：\n"
    "  y(t) = Σ wₖ* · sₖ(t)\n\n"
    "相位对齐条件：\n"
    "  ∠wₖ = −(2πfcτₖ + φₖ)"
)
add_textbox(slide1, model_text,
            Inches(4.8), Inches(1.52), Inches(4.2), Inches(2.35),
            font_size=12, color=WHITE)

add_rect(slide1, Inches(4.65), Inches(4.1), Inches(4.5), Inches(2.85), MID_BLUE)
add_textbox(slide1, "相参增益退化律（核心公式）",
            Inches(4.8), Inches(4.15), Inches(4.2), Inches(0.38),
            font_size=14, bold=True, color=ACCENT_GOLD)
gain_text = (
    "实际合并增益：\n"
    "  G = N² · sinc²(σ_φ_total)\n\n"
    "联合相位误差：\n"
    "  σ²_φ = σ²_time·(2πfc)²\n"
    "        + σ²_phase + σ²_freq · T²\n\n"
    "  σ_φ < π/8  →  G > 0.90 · N²  (保持90%效能)\n"
    "  σ_φ = π/4  →  G ≈ 0.64 · N²  (退化36%)"
)
add_textbox(slide1, gain_text,
            Inches(4.8), Inches(4.53), Inches(4.2), Inches(2.35),
            font_size=12, color=WHITE)

# ── 右侧：增益对比说明 + 框图 ──
add_rect(slide1, Inches(9.4), Inches(1.05), Inches(3.7), Inches(1.55), MID_BLUE)
add_textbox(slide1, "增益对比",
            Inches(9.55), Inches(1.1), Inches(3.4), Inches(0.35),
            font_size=13, bold=True, color=ACCENT_GOLD)
compare_text = (
    "相参合并：SNR ∝ N²\n"
    "非相参合并：SNR ∝ N\n"
    "相参优势：额外 N 倍增益\n"
    "（10·lgN dB 提升）"
)
add_textbox(slide1, compare_text,
            Inches(9.55), Inches(1.45), Inches(3.4), Inches(1.1),
            font_size=11, color=WHITE)

# 简化系统框图（文字版）
add_rect(slide1, Inches(9.4), Inches(2.75), Inches(3.7), Inches(4.2), MID_BLUE)
add_textbox(slide1, "系统框图",
            Inches(9.55), Inches(2.8), Inches(3.4), Inches(0.35),
            font_size=13, bold=True, color=ACCENT_GOLD)

boxes = ["节点 1 接收", "节点 2 接收", "节点 N 接收"]
for i, b in enumerate(boxes):
    add_rect(slide1, Inches(9.5), Inches(3.22) + i*Inches(0.75),
             Inches(1.5), Inches(0.55), RGBColor(0x10,0x5B,0x9E),
             line_color=ACCENT_GOLD, line_width=Pt(1))
    add_textbox(slide1, b,
                Inches(9.5), Inches(3.22) + i*Inches(0.75),
                Inches(1.5), Inches(0.55), font_size=10, align=PP_ALIGN.CENTER)

add_textbox(slide1, "↓  ↓  ↓",
            Inches(9.5), Inches(5.49), Inches(1.5), Inches(0.3),
            font_size=11, color=ACCENT_GOLD, align=PP_ALIGN.CENTER)
add_rect(slide1, Inches(9.5), Inches(5.79),
         Inches(1.5), Inches(0.55), RGBColor(0x0E,0x6E,0x6E),
         line_color=ACCENT_GOLD, line_width=Pt(1))
add_textbox(slide1, "相位对齐\n（补偿器）",
            Inches(9.5), Inches(5.79),
            Inches(1.5), Inches(0.55), font_size=10, align=PP_ALIGN.CENTER)

add_textbox(slide1, "→",
            Inches(11.05), Inches(5.79), Inches(0.35), Inches(0.55),
            font_size=18, color=ACCENT_GOLD, align=PP_ALIGN.CENTER)
add_rect(slide1, Inches(11.4), Inches(5.79),
         Inches(1.5), Inches(0.55), RGBColor(0x7A,0x3A,0x10),
         line_color=ACCENT_GOLD, line_width=Pt(1))
add_textbox(slide1, "MRC 合并\n→ 增益 N²",
            Inches(11.4), Inches(5.79),
            Inches(1.5), Inches(0.55), font_size=10, align=PP_ALIGN.CENTER)

# ── 底部注释 ──
add_textbox(slide1,
            "★ 空海特殊挑战：载体运动导致基线动态变化，多普勒快变，四大误差耦合非线性退化，同步维持难度远高于地基系统",
            Inches(0.25), Inches(7.1), Inches(12.8), Inches(0.35),
            font_size=10, italic=True, color=LIGHT_GREY)


# ══════════════════════════════════════════════════════════
#  第 2 页：相参合成方法设计
# ══════════════════════════════════════════════════════════
slide2 = prs.slides.add_slide(blank_layout)
set_bg(slide2, DARK_BLUE)

# ── 顶部标题栏 ──
add_rect(slide2, Inches(0), Inches(0), SW, Inches(0.85), MID_BLUE)
add_textbox(slide2,
            "空海平台多节点协同探测  |  高效相参合成方法设计",
            Inches(0.3), Inches(0.08), Inches(12), Inches(0.65),
            font_size=22, bold=True, color=ACCENT_GOLD, align=PP_ALIGN.LEFT)

# ── 流水线框图（顶部横排）──
pipeline_steps = [
    ("① 原始回波\n输入", RGBColor(0x10,0x5B,0x9E)),
    ("② 时延估计\n与对齐",  RGBColor(0x0E,0x6E,0x6E)),
    ("③ 相位估计\n与校正",  RGBColor(0x4A,0x23,0x7A)),
    ("④ 自适应\n权值计算",  RGBColor(0x7A,0x3A,0x10)),
    ("⑤ 相参合并\n输出",    RGBColor(0x1B,0x6B,0x2E)),
    ("⑥ 检测\n处理",        RGBColor(0x6B,0x1B,0x1B)),
]
pipe_w = Inches(1.75)
pipe_h = Inches(0.72)
for i, (label, col) in enumerate(pipeline_steps):
    lx = Inches(0.22) + i * (pipe_w + Inches(0.18))
    add_rect(slide2, lx, Inches(1.0), pipe_w, pipe_h, col,
             line_color=ACCENT_GOLD, line_width=Pt(1))
    add_textbox(slide2, label, lx, Inches(1.0), pipe_w, pipe_h,
                font_size=11, bold=True, align=PP_ALIGN.CENTER)
    if i < len(pipeline_steps) - 1:
        add_textbox(slide2, "→",
                    lx + pipe_w, Inches(1.0), Inches(0.18), pipe_h,
                    font_size=14, color=ACCENT_GOLD, align=PP_ALIGN.CENTER)

# ── 四列方法详解 ──
sections = [
    {
        "title": "① 时延对齐方法",
        "color": RGBColor(0x10, 0x5B, 0x9E),
        "items": [
            "北斗/GPS双频授时",
            "硬件同步精度 < 10 ns",
            "互相关峰值精估：精度≈1/BW",
            "INS/DVL动态时延预测",
            "补偿载体机动引起时延抖动",
            "目标：Δτ < λ/(4c)",
        ]
    },
    {
        "title": "② 相位校正方法",
        "color": RGBColor(0x0E, 0x6E, 0x6E),
        "items": [
            "导频信号法（有导频）：",
            "  • 周期发送相位参考信号",
            "  • LS拟合载波相位漂移",
            "  • 实时反馈补偿",
            "目标回波自校准（无导频）：",
            "  • 利用多节点回波相位差",
            "  • 反解并补偿相位误差",
            "EKF/PLL 联合跟踪框架",
        ]
    },
    {
        "title": "③ 自适应权值计算",
        "color": RGBColor(0x4A, 0x23, 0x7A),
        "items": [
            "最大比合并（MRC）：",
            "  wₖ = SNRₖ / Σ SNRᵢ",
            "稳健加权：",
            "  • 相位误差大的节点自动降权",
            "  • 防止[相参污染]",
            "分层相参策略：",
            "  • 同步质量相近节点组内相参",
            "  • 组间非相参合并",
        ]
    },
    {
        "title": "④ 空海特殊设计",
        "color": RGBColor(0x7A, 0x3A, 0x10),
        "items": [
            "海面多径处理：",
            "  • 分离直射/镜面反射分量",
            "  • 分量分离后再相参合并",
            "载体姿态补偿：",
            "  • IMU实时修正横摇/纵摇",
            "  • Δφ=(2π/λ)·d·sin(Δθ)",
            "低截获设计：",
            "  • 相参合并在数字域完成",
        ]
    },
]

col_w = Inches(3.15)
for ci, sec in enumerate(sections):
    lx = Inches(0.22) + ci * (col_w + Inches(0.12))
    # 列标题块
    add_rect(slide2, lx, Inches(1.9), col_w, Inches(0.42), sec["color"],
             line_color=ACCENT_GOLD, line_width=Pt(1))
    add_textbox(slide2, sec["title"], lx, Inches(1.9), col_w, Inches(0.42),
                font_size=12, bold=True, color=ACCENT_GOLD, align=PP_ALIGN.CENTER)
    # 内容背景
    add_rect(slide2, lx, Inches(2.35), col_w, Inches(2.85),
             RGBColor(0x12, 0x2A, 0x4A))
    body_text = "\n".join(f"{'  ' if item.startswith('  ') else '• '}{item.lstrip()}"
                          if not item.startswith('  ') else item
                          for item in sec["items"])
    # 简单逐行写
    y_offset = Inches(2.38)
    for item in sec["items"]:
        indent = item.startswith("  ")
        prefix = "   " if indent else "• "
        add_textbox(slide2, prefix + item.strip(),
                    lx + Inches(0.08), y_offset, col_w - Inches(0.16), Inches(0.32),
                    font_size=10,
                    color=LIGHT_GREY if indent else WHITE)
        y_offset += Inches(0.32)

# ── 底部性能对比表 ──
add_textbox(slide2, "关键性能指标对比",
            Inches(0.22), Inches(5.28), Inches(5), Inches(0.35),
            font_size=13, bold=True, color=ACCENT_GOLD)

tbl = add_table(slide2, 4, 4, Inches(0.22), Inches(5.65),
                Inches(12.85), Inches(1.65))

headers = ["方法类型", "同步精度要求", "算法延迟", "适用场景"]
rows_data = [
    ["导频相参合并",   "时延 < 10 ns",  "低（< 1 ms）",   "协同组网、慢机动平台"],
    ["回波自校准合并", "时延 < 50 ns",  "中（1 ~ 10 ms）", "无导频、通信受限场景"],
    ["分层相参合并",   "时延 < 100 ns", "低",              "节点数多、同步质量异构"],
]

for ci, h in enumerate(headers):
    cell_style(tbl.cell(0, ci), h, font_size=12, bold=True, bg_color=TABLE_HEADER,
               font_color=ACCENT_GOLD)

row_bg = [RGBColor(0x12,0x2A,0x4A), RGBColor(0x0D,0x22,0x3C), RGBColor(0x12,0x2A,0x4A)]
for ri, row in enumerate(rows_data):
    for ci, val in enumerate(row):
        cell_style(tbl.cell(ri+1, ci), val, font_size=11,
                   bg_color=row_bg[ri], font_color=WHITE,
                   align=PP_ALIGN.LEFT if ci == 0 or ci == 3 else PP_ALIGN.CENTER)

# ── 底部注释 ──
add_textbox(slide2,
            "★ 核心权衡：空海高动态平台要求在「估计精度 — 计算实时性 — 通信开销」三者间取得最优平衡",
            Inches(0.22), Inches(7.12), Inches(12.85), Inches(0.33),
            font_size=10, italic=True, color=LIGHT_GREY)


# ── 保存文件 ──────────────────────────────────────────────
output_path = "/home/runner/work/project/project/coherent_synthesis_slides.pptx"
prs.save(output_path)
print(f"✅ 幻灯片已生成：{output_path}")
