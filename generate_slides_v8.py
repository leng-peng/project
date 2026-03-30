"""
generate_slides_v8.py
─────────────────────────────────────────────────────────────────────────
多节点信号级高效相参合成 · 研究内容与技术路线（v8）

  单页 PPT，布局与 v6/v7 保持一致（相同调色板、相同工具函数）：
    ① 顶部全宽"研究思路"流程带（科学问题→研究内容→关键技术→预期成果）
    ② 主内容区 · 左列「三项研究内容」（绿色/青色系，3 个子卡片）
       - 研究内容一：增益退化端到端量化模型
       - 研究内容二：高动态时变相位误差建模与实时估计
       - 研究内容三：质量自适应MRC合并与层次化处理框架
    ③ 主内容区 · 右列「关键技术方法与预期指标」（蓝/紫色系，3 个子卡片）
       - 方法①：误差传播链路建模与效能边界推导
       - 方法②：非平稳相位在线估计（EKF/自校准）
       - 方法③：分层相参策略与实时计算优化

运行：
    python generate_slides_v8.py
输出：
    coherent_synthesis_slides_v8.pptx
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

# 主题色 — 研究内容（绿色/青色系）
RC_HDR   = RGBColor(0x0A, 0x5C, 0x36)   # 深绿（区标题）
RC_SUB1  = RGBColor(0x15, 0x7A, 0x35)   # 子卡片1（绿）
RC_SUB2  = RGBColor(0x00, 0x7C, 0x78)   # 子卡片2（青）
RC_SUB3  = RGBColor(0x0E, 0x6B, 0x5E)   # 子卡片3（深青）

# 主题色 — 关键技术（蓝/紫色系）
KT_HDR   = RGBColor(0x0D, 0x2A, 0x6E)   # 深蓝（区标题）
KT_SUB1  = RGBColor(0x1F, 0x6B, 0xC4)   # 子卡片1（蓝）
KT_SUB2  = RGBColor(0x6A, 0x2B, 0xB5)   # 子卡片2（紫）
KT_SUB3  = RGBColor(0xB5, 0x4A, 0x00)   # 子卡片3（橙，预期指标）

# 流程带配色（4步）
FLOW_COLS = [
    RGBColor(0xB5, 0x4A, 0x00),   # 科学问题（橙）
    RGBColor(0x0A, 0x5C, 0x36),   # 研究内容（深绿）
    RGBColor(0x1F, 0x6B, 0xC4),   # 关键技术（蓝）
    RGBColor(0x0D, 0x2A, 0x6E),   # 预期成果（深蓝）
]

COL_NAVY   = RGBColor(0x0D, 0x2A, 0x6E)
COL_GREEN  = RGBColor(0x0A, 0x5C, 0x36)

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
#  研究内容与技术路线页
#  ┌─────────────────────────────────────────────────────────────────┐
#  │  TITLE BAR                                                      │
#  ├─────────────────────────────────────────────────────────────────┤
#  │  研究思路流程带（科学问题→研究内容→关键技术→预期成果）          │
#  ├──────────────────────────┬──────────────────────────────────────┤
#  │  【三项研究内容】(绿色)  │  【关键技术方法与预期指标】(蓝/紫)  │
#  │  内容一：增益退化量化    │  方法①：误差传播链路建模             │
#  │  内容二：时变相位建模    │  方法②：EKF/自校准在线估计           │
#  │  内容三：自适应MRC框架  │  方法③：分层相参+实时计算优化       │
#  ├─────────────────────────────────────────────────────────────────┤
#  │  FOOTER                                                         │
#  └─────────────────────────────────────────────────────────────────┘
# ╚══════════════════════════════════════════════════════════════════════╝

s1 = prs.slides.add_slide(blank)
set_bg(s1, BG)
title_bar(s1,
          "多节点信号级高效相参合成 · 研究内容与技术路线",
          "Research Content & Technical Approach — From Scientific Problems to Engineering Solutions",
          1)

# ── 顶部：研究思路流程带 ─────────────────────────────────────────────
FLOW_TOP = CT
FLOW_H   = Inches(0.80)

flow_steps = [
    ("科学问题\n增益律/融合条件/CRLB",    FLOW_COLS[0]),
    ("研究内容\n三项核心研究任务",          FLOW_COLS[1]),
    ("关键技术\n误差建模/相位估计/MRC",    FLOW_COLS[2]),
    ("预期成果\n理论突破 + 工程验证",       FLOW_COLS[3]),
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

# 在"研究内容"和"关键技术"下方加竖向指示线
for idx in [1, 2]:
    hx = MARGIN + idx * (step_w + ARR_W) + step_w / 2
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

SEC_HDR_H = Inches(0.42)

# ── 左区大标题（绿色：研究内容）──
rect(s1, MARGIN + Inches(0.04), MAIN_TOP + Inches(0.04),
     L_W, SEC_HDR_H, RGBColor(0xBB, 0xCC, 0xBB))
rect(s1, MARGIN, MAIN_TOP, L_W, SEC_HDR_H, RC_HDR)
rect(s1, MARGIN, MAIN_TOP, Inches(0.10), SEC_HDR_H,
     RGBColor(0x80, 0xFF, 0xA0))   # 亮绿左侧强调条
tb(s1, "📋",
   MARGIN + Inches(0.14), MAIN_TOP + Inches(0.05),
   Inches(0.30), SEC_HDR_H - Inches(0.08),
   sz=14, bold=True, color=WHITE)
tb(s1, "三项研究内容   Research Objectives",
   MARGIN + Inches(0.50), MAIN_TOP + Inches(0.06),
   L_W - Inches(0.56), Inches(0.26),
   sz=13, bold=True, color=WHITE)
tb(s1, "面向三个科学问题，分别设计可量化验证的研究任务",
   MARGIN + Inches(0.50), MAIN_TOP + Inches(0.26),
   L_W - Inches(0.56), Inches(0.14),
   sz=8.5, color=RGBColor(0xC0, 0xFF, 0xD0))

# 左区三个子卡片
SUB_GAP = Inches(0.10)
sub_top = MAIN_TOP + SEC_HDR_H + Inches(0.10)
sub_h   = (MAIN_H - SEC_HDR_H - Inches(0.10) - 2 * SUB_GAP) / 3

rc_cards = [
    {
        "accent": RC_SUB1,
        "badge": "一",
        "title": "增益退化端到端量化模型（对应科学问题A）",
        "items": [
            ("建立时延误差σ_τ / 相位误差σ_φ / 频偏σ_f → 合并增益G的统一传播链路", False),
            ("G_eff(σ_τ, σ_φ, σ_f) = G · exp(−σ_φ²) · sinc²(σ_τ·B) · sinc²(σ_f·Nc·PRI)", True),
            ('量化"效能90%保持区"：σ_φ < π/8, σ_τ < 0.5 ns, σ_f < 100 Hz', False),
            ("在N=2,4,6节点下仿真验证退化曲线，给出工程容差设计准则", False),
        ]
    },
    {
        "accent": RC_SUB2,
        "badge": "二",
        "title": "高动态时变相位误差的建模与实时估计",
        "items": [
            ("建立非平稳相位模型：Δφ_k(t) = φ_delay(t) + φ_osc(t) + φ_Doppler(t)", True),
            ("无导频自校准：基于回波互相关 + EKF 实时跟踪时变相位漂移率", False),
            ("推导最小可辨识帧长 N_c,min 与 SNR_k 的联合约束（CRLB下界）", False),
            ("设计自适应帧长选择准则：在估计精度与计算延迟间动态权衡", False),
        ]
    },
    {
        "accent": RC_SUB3,
        "badge": "三",
        "title": "质量自适应MRC合并与层次化处理框架",
        "items": [
            ("MRC权值：w_k = SNR_k · coh_k · ph_cons_k，三因子联合在线估计", True),
            ("稳健机制：相位估计失败时自动降权，避免合并输出突变", False),
            ("分层策略：强目标SIC对消 → 弱目标三MTI检测 → 统一MRC合并", False),
            ("计算优化：子带分解 + 稀疏采样，将O(N·Nc·Nsc·lgNc)降低一量级", False),
        ]
    },
]

for ci, dc in enumerate(rc_cards):
    cy = sub_top + ci * (sub_h + SUB_GAP)
    sub_card(s1, MARGIN, cy, L_W, sub_h,
             dc["accent"], dc["title"], dc["items"],
             badge=dc["badge"])

# ── 右区大标题（蓝/紫色：关键技术与预期指标）──
rect(s1, R_X + Inches(0.04), MAIN_TOP + Inches(0.04),
     R_W, SEC_HDR_H, RGBColor(0xC0, 0xC8, 0xE0))
rect(s1, R_X, MAIN_TOP, R_W, SEC_HDR_H, KT_HDR)
rect(s1, R_X, MAIN_TOP, Inches(0.10), SEC_HDR_H,
     RGBColor(0x80, 0xB8, 0xFF))
tb(s1, "⚙",
   R_X + Inches(0.14), MAIN_TOP + Inches(0.05),
   Inches(0.30), SEC_HDR_H - Inches(0.08),
   sz=14, bold=True, color=WHITE)
tb(s1, "关键技术方法与预期指标",
   R_X + Inches(0.50), MAIN_TOP + Inches(0.06),
   R_W - Inches(0.56), Inches(0.26),
   sz=13, bold=True, color=WHITE)
tb(s1, "Key Methods & Expected Technical Metrics — Quantifiable Validation Targets",
   R_X + Inches(0.50), MAIN_TOP + Inches(0.26),
   R_W - Inches(0.56), Inches(0.14),
   sz=8.5, color=RGBColor(0xB0, 0xD8, 0xFF))

# 右区三个子卡片
kt_cards = [
    {
        "accent": KT_SUB1,
        "badge": "①",
        "title": "误差传播链路建模与效能边界推导",
        "items": [
            ("理论工具：联合误差传播理论 + 贝叶斯效能下界（BCRLB）", False),
            ("输出：G_eff vs (σ_τ, σ_φ, σ_f) 三维效能图谱及工程容差曲线", False),
            ("预期指标：N=6时理想G=15.6 dB，误差容限内G_eff ≥ 14 dB（90%）", False),
            ("验证方式：Monte-Carlo仿真（1000次）+ MATLAB端到端链路仿真", False),
        ]
    },
    {
        "accent": KT_SUB2,
        "badge": "②",
        "title": "EKF/自校准在线相位估计算法",
        "items": [
            ("状态量：[Δφ_k, φ_dot_k]，过程噪声由平台运动学参数驱动", True),
            ("量测更新：互相关亚像素峰值提供Δτ观测，EKF步长 = 1 CPI", False),
            ("预期指标：SNR≥6 dB时，相位估计误差 σ_est < π/8（满足G_eff ≥ 90%G）", False),
            ("对比基准：无补偿 / 静态互相关补偿 / 本方法三方案对比", False),
        ]
    },
    {
        "accent": KT_SUB3,
        "badge": "③",
        "title": "分层相参策略与实时性预期指标",
        "items": [
            ("Round1：无MTI强目标检测（CFAR>12dB）→ SIC对消 → 弱目标释放", False),
            ("Round2：三MTI并行（SLOW/FAST/VFAST）+ MRC → CFAR>6dB统一输出", False),
            ("预期处理延迟：两轮合计 < 9 ms（Nc=64, Nsc=1024, N=6链路）",        False),
            ("近似优化：子带分解压缩至 < 5 ms；精度损失 < 0.3 dB（仿真验证）",    True),
        ]
    },
]

for ci, kc in enumerate(kt_cards):
    cy = sub_top + ci * (sub_h + SUB_GAP)
    sub_card(s1, R_X, cy, R_W, sub_h,
             kc["accent"], kc["title"], kc["items"],
             badge=kc["badge"])

footer(s1,
       "研究路径：量化退化律（内容一）→ 实时相位估计（内容二）→ 自适应MRC框架（内容三）"
       "——三项研究形成完整闭环，最终实现 N=6 节点下 G_eff ≥ 14 dB、延迟 < 9 ms 的工程目标")


# ── 保存 ───────────────────────────────────────────────────────────
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                   "coherent_synthesis_slides_v8.pptx")
prs.save(OUT)
print(f"✅  已生成：{OUT}")
