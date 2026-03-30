"""
generate_framework_v9.py
────────────────────────────────────────────────────────────────────────
多节点高效相参合成总体研究框架图（PNG 格式）

布局（4行 × 3列 主体 + 左侧层次标签 + 底部高效核心栏）：
  ┌───────────────────────────────────────────────────────────────┐
  │                  总标题                                        │
  ├──────────┬──────────────┬──────────────┬──────────────────────┤
  │  层次    │  地基平台    │  空中平台    │  空海平台            │
  ├──────────┼──────────────┼──────────────┼──────────────────────┤
  │ 背景需求 │  ...         │  ...         │  ...                 │
  │          │      ↓       │      ↓       │      ↓               │
  │ 方法原理 │  ...         │  ...         │  ...                 │
  │          │      ↓       │      ↓       │      ↓               │
  │ 模型流程 │  ...         │  ...         │  ...                 │
  │          │      ↓       │      ↓       │      ↓               │
  │ 仿真实现 │  ...         │  ...         │  ...                 │
  ├──────────┴──────────────┴──────────────┴──────────────────────┤
  │              高效相参合成核心指标（共性底栏）                   │
  └───────────────────────────────────────────────────────────────┘

运行：
    python generate_framework_v9.py
输出：
    coherent_framework_v9.png
"""

import matplotlib
matplotlib.use('Agg')

import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import matplotlib.font_manager as fm
import os

# ─── 字体 ──────────────────────────────────────────────────────────────
FONT_REGULAR = '/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc'
FONT_BOLD    = '/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc'
fp_r = fm.FontProperties(fname=FONT_REGULAR) if os.path.exists(FONT_REGULAR) else fm.FontProperties()
fp_b = fm.FontProperties(fname=FONT_BOLD)    if os.path.exists(FONT_BOLD)    else fm.FontProperties(weight='bold')

# ─── 调色板 ────────────────────────────────────────────────────────────
C_BG        = '#F4F7FC'    # 背景
C_TITLE     = '#1F3864'    # 主标题栏深蓝
C_TITLE2    = '#2E74B5'    # 次标题蓝

# 地基平台 — 绿色系
C_GND_H     = '#1A5C38'   # 深绿（列头）
C_GND_L1    = '#2E7D52'   # 需求行
C_GND_L2    = '#3A9966'   # 原理行
C_GND_L3    = '#4BAE7E'   # 流程行
C_GND_L4    = '#6DC49A'   # 仿真行
C_GND_BG    = '#E8F5EE'   # 内容格背景

# 空中平台 — 蓝色系
C_AIR_H     = '#1435A0'   # 深蓝（列头）
C_AIR_L1    = '#1F56C8'
C_AIR_L2    = '#2D6ED8'
C_AIR_L3    = '#3C82E0'
C_AIR_L4    = '#609BE8'
C_AIR_BG    = '#E8EFFE'

# 空海平台 — 橙色系
C_SEA_H     = '#8B2E00'   # 深橙（列头）
C_SEA_L1    = '#B54200'
C_SEA_L2    = '#CC5810'
C_SEA_L3    = '#E07020'
C_SEA_L4    = '#F09040'
C_SEA_BG    = '#FFF0E5'

# 底栏 — 紫色
C_BOT_H     = '#3B0E6E'
C_BOT_BG    = '#EDE8F8'

C_ARROW     = '#555577'
C_WHITE     = '#FFFFFF'
C_TEXT_D    = '#1A1A2E'   # 深色正文
C_LAYER_BG  = '#D6E4F7'   # 层次标签背景

# ─── 画布 ──────────────────────────────────────────────────────────────
FIG_W, FIG_H = 22, 16     # inches，宽屏

fig, ax = plt.subplots(figsize=(FIG_W, FIG_H))
ax.set_xlim(0, FIG_W)
ax.set_ylim(0, FIG_H)
ax.axis('off')
fig.patch.set_facecolor(C_BG)
ax.set_facecolor(C_BG)


# ─── 工具函数 ──────────────────────────────────────────────────────────
def rbox(ax, x, y, w, h, fc, ec='none', lw=0, radius=0.18, zorder=2):
    box = mpatches.FancyBboxPatch(
        (x, y), w, h,
        boxstyle=f'round,pad=0,rounding_size={radius}',
        fc=fc, ec=ec, lw=lw, zorder=zorder,
        transform=ax.transData, clip_on=False
    )
    ax.add_patch(box)
    return box


def text(ax, x, y, s, fp, sz, color, ha='center', va='center', zorder=3):
    ax.text(x, y, s, fontproperties=fp, fontsize=sz,
            color=color, ha=ha, va=va, zorder=zorder,
            transform=ax.transData, clip_on=False)


def arrow_down(ax, x, y_top, length=0.30):
    ax.annotate('', xy=(x, y_top - length), xytext=(x, y_top),
                arrowprops=dict(arrowstyle='->', color=C_ARROW,
                                lw=1.8, mutation_scale=14),
                zorder=4)


def multiline_text(ax, cx, cy, lines, fp, sz, color, line_gap=0.17,
                   ha='center', va='center'):
    """在 cy 为中心处，逐行绘制 lines 列表"""
    n = len(lines)
    top_y = cy + (n - 1) * line_gap / 2
    for i, ln in enumerate(lines):
        text(ax, cx, top_y - i * line_gap, ln, fp, sz, color, ha=ha, va=va)


# ══════════════════════════════════════════════════════════════════════
# 1.  主标题栏
# ══════════════════════════════════════════════════════════════════════
TITLE_Y  = FIG_H - 1.00
TITLE_H  = 0.90
rbox(ax, 0.0, TITLE_Y, FIG_W, TITLE_H, C_TITLE, radius=0.1, zorder=1)
rbox(ax, 0.0, TITLE_Y, 0.22, TITLE_H, '#0D1E4A', radius=0.05, zorder=2)

text(ax, FIG_W / 2, TITLE_Y + TITLE_H * 0.60,
     '多节点高效相参合成总体研究框架图',
     fp_b, 26, C_WHITE)
text(ax, FIG_W / 2, TITLE_Y + TITLE_H * 0.22,
     'Overall Research Framework for Efficient Multi-Node Coherent Synthesis  ·  地基 / 空中 / 空海 三大平台',
     fp_r, 12, '#B0CCF0')

# ══════════════════════════════════════════════════════════════════════
# 2.  布局常量
# ══════════════════════════════════════════════════════════════════════
MARGIN   = 0.30
LABEL_W  = 1.30          # 左侧层次标签宽度
COL_GAP  = 0.20          # 列间隙
BOT_H    = 0.75          # 底部共性栏高度
BOT_Y    = 0.20          # 底部栏 y

CONTENT_TOP = TITLE_Y - 0.20       # 内容区顶部
CONTENT_BOT = BOT_Y + BOT_H + 0.18 # 内容区底部

CONTENT_H   = CONTENT_TOP - CONTENT_BOT   # 内容区总高
HEADER_H    = 0.62                 # 平台列头高度
ARROW_GAP   = 0.30                 # 行间箭头区高度
N_ROWS      = 4
BODY_H      = (CONTENT_H - HEADER_H - ARROW_GAP * (N_ROWS - 1)) / N_ROWS

# 三列
COL_X0 = MARGIN + LABEL_W + COL_GAP
AVAIL_W = FIG_W - MARGIN - COL_X0 - MARGIN
COL_W   = (AVAIL_W - COL_GAP * 2) / 3

col_xs = [COL_X0,
          COL_X0 + COL_W + COL_GAP,
          COL_X0 + 2 * (COL_W + COL_GAP)]
col_cx = [x + COL_W / 2 for x in col_xs]

# 层次行（从上到下）
row_labels  = ['背景\n需求', '方法\n原理', '模型\n流程', '仿真\n实现']
row_titles  = ['背景需求', '方法原理（高效相参）', '模型流程', '仿真实现']

# 从上往下计算每行顶部 y 坐标
row_ys = []
y_cur = CONTENT_TOP - HEADER_H
for i in range(N_ROWS):
    row_ys.append(y_cur)
    y_cur -= BODY_H
    if i < N_ROWS - 1:
        y_cur -= ARROW_GAP    # 箭头间隔

# ══════════════════════════════════════════════════════════════════════
# 3.  平台列头
# ══════════════════════════════════════════════════════════════════════
plat_cfgs = [
    ('地基平台\nGround-based',  C_GND_H,  [C_GND_L1, C_GND_L2, C_GND_L3, C_GND_L4], C_GND_BG),
    ('空中平台\nAirborne',      C_AIR_H,  [C_AIR_L1, C_AIR_L2, C_AIR_L3, C_AIR_L4], C_AIR_BG),
    ('空海平台\nAir-Sea',       C_SEA_H,  [C_SEA_L1, C_SEA_L2, C_SEA_L3, C_SEA_L4], C_SEA_BG),
]

HEADER_TOP = CONTENT_TOP - HEADER_H
for i, (title, hdr_col, row_cols, bg_col) in enumerate(plat_cfgs):
    cx = col_cx[i]
    x0 = col_xs[i]
    # 列头背景
    rbox(ax, x0, HEADER_TOP, COL_W, HEADER_H, hdr_col, radius=0.20, zorder=2)
    lines = title.split('\n')
    text(ax, cx, HEADER_TOP + HEADER_H * 0.60, lines[0], fp_b, 16, C_WHITE)
    text(ax, cx, HEADER_TOP + HEADER_H * 0.22, lines[1], fp_r, 10, '#DDEEFF')

# ══════════════════════════════════════════════════════════════════════
# 4.  左侧层次标签
# ══════════════════════════════════════════════════════════════════════
LAYER_COLS = ['#2C5E8A', '#1A5C38', '#7A3B0E', '#3B0E6E']
for i, (lbl, lyr_col) in enumerate(zip(row_labels, LAYER_COLS)):
    y0 = row_ys[i]
    cy = y0 - BODY_H / 2
    rbox(ax, MARGIN, y0 - BODY_H, LABEL_W, BODY_H, lyr_col, radius=0.18, zorder=2)
    for j, line_str in enumerate(lbl.split('\n')):
        off = 0.16 if j == 0 else -0.16
        text(ax, MARGIN + LABEL_W / 2, cy + off, line_str, fp_b, 15, C_WHITE)

# ══════════════════════════════════════════════════════════════════════
# 5.  内容格数据
# ══════════════════════════════════════════════════════════════════════
# [row][col] => list of text lines
CELL_DATA = [
    # ── 行0：背景需求 ──────────────────────────────────────────────
    [
        # 地基平台
        ['■ 多固定站远程组网探测',
         '■ 时间/频率/相位同步约束',
         '■ 高增益低虚警检测需求',
         '■ 稳定基线·低机动性'],
        # 空中平台
        ['■ 机载编队协同作战',
         '■ 高速运动·基线时变',
         '■ 平台振动引入相位抖动',
         '■ 强杂波背景·实时处理'],
        # 空海平台
        ['■ 舰机异构协同探测',
         '■ 双/多基地构型变化',
         '■ 强海杂波·多径效应',
         '■ 异构量测配准需求'],
    ],
    # ── 行1：方法原理 ──────────────────────────────────────────────
    [
        # 地基平台
        ['◆ 相参合成增益理论模型',
         '◆ 最大比合并(MRC)加权',
         '◆ 增益退化律量化分析',
         '◆ 相位误差传播链路建模'],
        # 空中平台
        ['◆ 时变相位动态建模',
         '◆ EKF在线相位估计与预测',
         '◆ 运动补偿+多普勒配准',
         '◆ 自适应加权融合机制'],
        # 空海平台
        ['◆ 双基地信号等效转换',
         '◆ 异构量测统一域配准',
         '◆ 分层信号级融合策略',
         '◆ 海面散射相位特性建模'],
    ],
    # ── 行2：模型流程 ──────────────────────────────────────────────
    [
        # 地基平台
        ['① 多站回波同步采集',
         '② 脉冲压缩 → 距离对齐',
         '③ 相位校准 → 补偿',
         '④ MRC加权合并 → 检测'],
        # 空中平台
        ['① 平台运动状态估计',
         '② 多普勒域配准/补偿',
         '③ 时变相位预测更新',
         '④ 自适应融合 → 检测'],
        # 空海平台
        ['① 双基地RD图生成',
         '② 空间/时间坐标配准',
         '③ 统一域相参合并',
         '④ 联合CFAR检测'],
    ],
    # ── 行3：仿真实现 ──────────────────────────────────────────────
    [
        # 地基平台
        ['▶ 阵列相参增益蒙特卡洛验证',
         '▶ 相位误差灵敏度曲线',
         '▶ SINR提升量化(N节点律)',
         '▶ 同步精度对增益影响'],
        # 空中平台
        ['▶ 高动态相参合成增益验证',
         '▶ EKF相位估计均方误差',
         '▶ 与CRLB对比分析',
         '▶ 实时处理复杂度评估'],
        # 空海平台
        ['▶ 异构平台相参效能分析',
         '▶ 海杂波背景下相参增益',
         '▶ 配准误差对增益退化影响',
         '▶ 双基地探测概率曲线'],
    ],
]

# ══════════════════════════════════════════════════════════════════════
# 6.  绘制内容格 + 行间箭头
# ══════════════════════════════════════════════════════════════════════
for row_i in range(N_ROWS):
    y_top = row_ys[row_i]
    y_bot = y_top - BODY_H
    row_cy = y_bot + BODY_H / 2

    for col_i, (plat_title, hdr_col, row_cols, bg_col) in enumerate(plat_cfgs):
        x0 = col_xs[col_i]
        cx = col_cx[col_i]
        accent = row_cols[row_i]

        # 格子背景（浅色）+ 左侧强调色竖条
        rbox(ax, x0, y_bot, COL_W, BODY_H, bg_col, ec=accent, lw=1.2, radius=0.14, zorder=2)
        rbox(ax, x0, y_bot, 0.12, BODY_H, accent, radius=0.06, zorder=3)

        # 文本
        lines = CELL_DATA[row_i][col_i]
        n = len(lines)
        v_step = min(0.20, (BODY_H - 0.20) / max(n, 1))
        y_start = y_bot + BODY_H / 2 + v_step * (n - 1) / 2
        for k, ln in enumerate(lines):
            text(ax, cx + 0.04, y_start - k * v_step, ln,
                 fp_r, 9.5, C_TEXT_D, ha='center', va='center')

    # 行间箭头（第0~2行之后）
    if row_i < N_ROWS - 1:
        arrow_y_top = y_bot   # 该行底部
        arrow_y_bot = row_ys[row_i + 1]   # 下一行顶部
        mid_y = (arrow_y_top + arrow_y_bot) / 2 + 0.03
        for cx in col_cx:
            arrow_down(ax, cx, arrow_y_top - 0.04, length=ARROW_GAP - 0.08)

# ══════════════════════════════════════════════════════════════════════
# 7.  底部：高效相参合成核心指标共性栏
# ══════════════════════════════════════════════════════════════════════
rbox(ax, MARGIN, BOT_Y, FIG_W - 2 * MARGIN, BOT_H, C_BOT_H, radius=0.15, zorder=2)

bot_items = [
    '高效性核心指标',
    '>> 相参增益 N^2（理论上限）',
    '>> 相位误差 sigma <= pi/8（3dB损失阈值）',
    '>> 计算复杂度 O(N*M)（线性可扩展）',
    '>> 数据率 >= 目标相干时间倒数',
    '>> 实时延迟 <= 雷达脉冲重复间隔',
]
xs = [MARGIN + 1.0 + i * (FIG_W - 2 * MARGIN - 2.0) / (len(bot_items) - 1) for i in range(len(bot_items))]
for i, (bx, item) in enumerate(zip(xs, bot_items)):
    fp = fp_b if i == 0 else fp_r
    sz = 11.5 if i == 0 else 9.8
    text(ax, bx, BOT_Y + BOT_H / 2, item, fp, sz, C_WHITE)
    if i == 0:
        # 分隔竖线
        ax.plot([bx + 1.3, bx + 1.3], [BOT_Y + 0.12, BOT_Y + BOT_H - 0.12],
                color='#9FAFE0', lw=1.2, zorder=3)

# ══════════════════════════════════════════════════════════════════════
# 8.  左侧：层次序号徽标
# ══════════════════════════════════════════════════════════════════════
layer_badges = ['①', '②', '③', '④']
badge_cols   = ['#2C5E8A', '#1A5C38', '#7A3B0E', '#3B0E6E']
for i in range(N_ROWS):
    y0 = row_ys[i]
    cy = y0 - BODY_H / 2
    badge_x = MARGIN + LABEL_W + 0.01
    circle = plt.Circle((badge_x - 0.28, cy + BODY_H / 2 - 0.22),
                         0.18, color=C_WHITE, zorder=5,
                         transform=ax.transData, clip_on=False)
    ax.add_patch(circle)
    text(ax, badge_x - 0.28, cy + BODY_H / 2 - 0.22,
         layer_badges[i], fp_b, 11, badge_cols[i])

# ══════════════════════════════════════════════════════════════════════
# 9.  右侧：平台标注装饰线
# ══════════════════════════════════════════════════════════════════════
for i, (plat_title, hdr_col, row_cols, bg_col) in enumerate(plat_cfgs):
    x0 = col_xs[i]
    x1 = x0 + COL_W
    bottom_y = row_ys[N_ROWS - 1] - BODY_H
    ax.plot([x0, x1], [bottom_y - 0.06, bottom_y - 0.06],
            color=hdr_col, lw=2.5, zorder=4, solid_capstyle='round')

# ══════════════════════════════════════════════════════════════════════
# 10. 连通整体的左侧纵向流程线
# ══════════════════════════════════════════════════════════════════════
flow_line_x = MARGIN + 0.12
flow_top = HEADER_TOP
flow_bot = BOT_Y + BOT_H
ax.annotate('', xy=(flow_line_x, flow_bot + 0.05),
            xytext=(flow_line_x, flow_top),
            arrowprops=dict(arrowstyle='->', color='#2E74B5',
                            lw=2.0, mutation_scale=16),
            zorder=4)
text(ax, flow_line_x - 0.05, (flow_top + flow_bot) / 2,
     '研\n究\n链\n路', fp_b, 9, '#2E74B5')

# ══════════════════════════════════════════════════════════════════════
# 保存
# ══════════════════════════════════════════════════════════════════════
OUT = '/home/runner/work/project/project/coherent_framework_v9.png'
plt.tight_layout(pad=0)
plt.savefig(OUT, dpi=150, bbox_inches='tight',
            facecolor=C_BG, edgecolor='none')
plt.close()
print(f'Saved: {OUT}')
