"""
generate_slides_v13.py
──────────────────────────────────────────────────────────────────────────
多节点信号级高效相参合成 · 研究内容与技术路线（v13）
基于 v8 内容，将文本公式替换为 MathType(OMML) 格式真实数学公式，
并对字体大小进行优化以提升可读性。

布局（16:9  13.33" × 7.50"）：
  ┌──────────────────────────────────────────────────────────────────┐
  │  标题栏                                                            │
  ├──────────────────────────────────────────────────────────────────┤
  │  研究思路流程带（科学问题→研究内容→关键技术→预期成果）              │
  ├──────────────────────────┬───────────────────────────────────────┤
  │  三项研究内容（绿色系）    │  关键技术方法与预期指标（蓝/紫色系）   │
  │  内容一 + OMML公式①      │  方法① 误差传播链路建模               │
  │  内容二 + OMML公式②      │  方法② EKF/自校准在线估计 + OMML公式④│
  │  内容三 + OMML公式③      │  方法③ 分层相参策略                   │
  ├──────────────────────────┴───────────────────────────────────────┤
  │  页脚                                                              │
  └──────────────────────────────────────────────────────────────────┘

OMML 公式（以 a14:m / m:oMath 嵌入到 PowerPoint 段落 XML）：
  ① G_eff(σ_τ, σ_φ, σ_f) = G · exp(−σ_φ²) · sinc²(σ_τ·B) · sinc²(σ_f·N_c·PRI)
  ② Δφ_k(t) = φ_delay(t) + φ_osc(t) + φ_Doppler(t)
  ③ w_k = SNR_k · coh_k · ph_cons_k
  ④ 状态量 [Δφ_k, φ̇_k]

运行：
    python generate_slides_v13.py
输出：
    coherent_synthesis_slides_v13.pptx
"""

import os
from lxml import etree

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# ━━━━━━━━━━━━━━━━━━━━━━ 调色板（与 v8 一致）━━━━━━━━━━━━━━━━━━━━━━━━
BG           = RGBColor(0xF4, 0xF7, 0xFC)
CARD_BG      = RGBColor(0xFF, 0xFF, 0xFF)
TITLE_BAR    = RGBColor(0x1F, 0x38, 0x64)
TITLE_BAR2   = RGBColor(0x2E, 0x74, 0xB5)
DARK_TEXT    = RGBColor(0x1A, 0x1A, 0x3A)
WHITE        = RGBColor(0xFF, 0xFF, 0xFF)
CARD_BORDER  = RGBColor(0xC8, 0xD8, 0xF0)
SEPARATOR    = RGBColor(0xC8, 0xD8, 0xF0)
FOOTER_BG    = RGBColor(0xE8, 0xF0, 0xFB)

RC_HDR   = RGBColor(0x0A, 0x5C, 0x36)
RC_SUB1  = RGBColor(0x15, 0x7A, 0x35)
RC_SUB2  = RGBColor(0x00, 0x7C, 0x78)
RC_SUB3  = RGBColor(0x0E, 0x6B, 0x5E)

KT_HDR   = RGBColor(0x0D, 0x2A, 0x6E)
KT_SUB1  = RGBColor(0x1F, 0x6B, 0xC4)
KT_SUB2  = RGBColor(0x6A, 0x2B, 0xB5)
KT_SUB3  = RGBColor(0xB5, 0x4A, 0x00)

FLOW_COLS = [
    RGBColor(0xB5, 0x4A, 0x00),
    RGBColor(0x0A, 0x5C, 0x36),
    RGBColor(0x1F, 0x6B, 0xC4),
    RGBColor(0x0D, 0x2A, 0x6E),
]

COL_NAVY  = RGBColor(0x0D, 0x2A, 0x6E)
COL_GREEN = RGBColor(0x0A, 0x5C, 0x36)

# 公式框颜色
MATH_BG   = RGBColor(0xE8, 0xF0, 0xFF)  # 淡蓝公式背景
MATH_BD   = RGBColor(0x2E, 0x74, 0xB5)  # 公式框边线

# ━━━━━━━━━━━━━━━━━━━━━━ 尺寸常量 ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
SW       = Inches(13.33)
SH       = Inches(7.50)
MARGIN   = Inches(0.32)
TITLE_H  = Inches(0.82)
CT       = TITLE_H + Inches(0.16)
FOOTER_Y = SH - Inches(0.50)
CB       = FOOTER_Y - Inches(0.04)
CW       = SW - 2 * MARGIN

# XML 命名空间
NS_A    = 'http://schemas.openxmlformats.org/drawingml/2006/main'
NS_A14  = 'http://schemas.microsoft.com/office/drawing/2010/main'
NS_MATH = 'http://schemas.openxmlformats.org/officeDocument/2006/math'


# ━━━━━━━━━━━━━━━━━━━━━━ OMML 公式定义 ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 以下为四个核心公式的 OMML XML 内容（m:oMath 内部片段）

# ① G_eff(σ_τ, σ_φ, σ_f) = G · exp(−σ_φ²) · sinc²(σ_τ·B) · sinc²(σ_f·Nc·PRI)
OMML_GEFF = """
<m:sSub xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <m:e><m:r><m:rPr><m:sty m:val="bi"/></m:rPr><m:t>G</m:t></m:r></m:e>
  <m:sub><m:r><m:t>eff</m:t></m:r></m:sub>
</m:sSub>
<m:d xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <m:dPr><m:ctrlPr/></m:dPr>
  <m:e>
    <m:sSub><m:e><m:r><m:t>&#x3C3;</m:t></m:r></m:e><m:sub><m:r><m:t>&#x3C4;</m:t></m:r></m:sub></m:sSub>
    <m:r><m:t xml:space="preserve">, </m:t></m:r>
    <m:sSub><m:e><m:r><m:t>&#x3C3;</m:t></m:r></m:e><m:sub><m:r><m:t>&#x3C6;</m:t></m:r></m:sub></m:sSub>
    <m:r><m:t xml:space="preserve">, </m:t></m:r>
    <m:sSub><m:e><m:r><m:t>&#x3C3;</m:t></m:r></m:e><m:sub><m:r><m:t>f</m:t></m:r></m:sub></m:sSub>
  </m:e>
</m:d>
<m:r xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <m:t xml:space="preserve"> = </m:t>
</m:r>
<m:r xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <m:rPr><m:sty m:val="bi"/></m:rPr>
  <m:t xml:space="preserve">G</m:t>
</m:r>
<m:r xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <m:t xml:space="preserve"> &#x22C5; </m:t>
</m:r>
<m:func xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <m:funcPr><m:ctrlPr/></m:funcPr>
  <m:fName><m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>exp</m:t></m:r></m:fName>
  <m:e>
    <m:d>
      <m:dPr><m:ctrlPr/></m:dPr>
      <m:e>
        <m:r><m:t>&#x2212;</m:t></m:r>
        <m:sSubSup>
          <m:e><m:r><m:t>&#x3C3;</m:t></m:r></m:e>
          <m:sub><m:r><m:t>&#x3C6;</m:t></m:r></m:sub>
          <m:sup><m:r><m:t>2</m:t></m:r></m:sup>
        </m:sSubSup>
      </m:e>
    </m:d>
  </m:e>
</m:func>
<m:r xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <m:t xml:space="preserve"> &#x22C5; </m:t>
</m:r>
<m:sSup xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <m:e>
    <m:func>
      <m:funcPr><m:ctrlPr/></m:funcPr>
      <m:fName><m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>sinc</m:t></m:r></m:fName>
      <m:e>
        <m:d>
          <m:dPr><m:ctrlPr/></m:dPr>
          <m:e>
            <m:sSub><m:e><m:r><m:t>&#x3C3;</m:t></m:r></m:e><m:sub><m:r><m:t>&#x3C4;</m:t></m:r></m:sub></m:sSub>
            <m:r><m:t xml:space="preserve">&#x22C5;B</m:t></m:r>
          </m:e>
        </m:d>
      </m:e>
    </m:func>
  </m:e>
  <m:sup><m:r><m:t>2</m:t></m:r></m:sup>
</m:sSup>
<m:r xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <m:t xml:space="preserve"> &#x22C5; </m:t>
</m:r>
<m:sSup xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <m:e>
    <m:func>
      <m:funcPr><m:ctrlPr/></m:funcPr>
      <m:fName><m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>sinc</m:t></m:r></m:fName>
      <m:e>
        <m:d>
          <m:dPr><m:ctrlPr/></m:dPr>
          <m:e>
            <m:sSub><m:e><m:r><m:t>&#x3C3;</m:t></m:r></m:e><m:sub><m:r><m:t>f</m:t></m:r></m:sub></m:sSub>
            <m:r><m:t xml:space="preserve">&#x22C5;</m:t></m:r>
            <m:sSub><m:e><m:r><m:t>N</m:t></m:r></m:e><m:sub><m:r><m:t>c</m:t></m:r></m:sub></m:sSub>
            <m:r><m:t xml:space="preserve">&#x22C5;PRI</m:t></m:r>
          </m:e>
        </m:d>
      </m:e>
    </m:func>
  </m:e>
  <m:sup><m:r><m:t>2</m:t></m:r></m:sup>
</m:sSup>
"""

# ② Δφ_k(t) = φ_delay(t) + φ_osc(t) + φ_Doppler(t)
OMML_PHASE = """
<m:sSub xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <m:e><m:r><m:rPr><m:sty m:val="bi"/></m:rPr><m:t>&#x394;&#x3C6;</m:t></m:r></m:e>
  <m:sub><m:r><m:t>k</m:t></m:r></m:sub>
</m:sSub>
<m:d xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <m:dPr><m:ctrlPr/></m:dPr>
  <m:e><m:r><m:rPr><m:sty m:val="i"/></m:rPr><m:t>t</m:t></m:r></m:e>
</m:d>
<m:r xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <m:t xml:space="preserve"> = </m:t>
</m:r>
<m:sSub xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <m:e><m:r><m:rPr><m:sty m:val="i"/></m:rPr><m:t>&#x3C6;</m:t></m:r></m:e>
  <m:sub><m:r><m:t>delay</m:t></m:r></m:sub>
</m:sSub>
<m:d xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <m:dPr><m:ctrlPr/></m:dPr>
  <m:e><m:r><m:rPr><m:sty m:val="i"/></m:rPr><m:t>t</m:t></m:r></m:e>
</m:d>
<m:r xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <m:t xml:space="preserve"> + </m:t>
</m:r>
<m:sSub xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <m:e><m:r><m:rPr><m:sty m:val="i"/></m:rPr><m:t>&#x3C6;</m:t></m:r></m:e>
  <m:sub><m:r><m:t>osc</m:t></m:r></m:sub>
</m:sSub>
<m:d xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <m:dPr><m:ctrlPr/></m:dPr>
  <m:e><m:r><m:rPr><m:sty m:val="i"/></m:rPr><m:t>t</m:t></m:r></m:e>
</m:d>
<m:r xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <m:t xml:space="preserve"> + </m:t>
</m:r>
<m:sSub xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <m:e><m:r><m:rPr><m:sty m:val="i"/></m:rPr><m:t>&#x3C6;</m:t></m:r></m:e>
  <m:sub><m:r><m:t>Doppler</m:t></m:r></m:sub>
</m:sSub>
<m:d xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <m:dPr><m:ctrlPr/></m:dPr>
  <m:e><m:r><m:rPr><m:sty m:val="i"/></m:rPr><m:t>t</m:t></m:r></m:e>
</m:d>
"""

# ③ w_k = SNR_k · coh_k · ph_cons_k
OMML_MRC = """
<m:sSub xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <m:e><m:r><m:rPr><m:sty m:val="bi"/></m:rPr><m:t>w</m:t></m:r></m:e>
  <m:sub><m:r><m:t>k</m:t></m:r></m:sub>
</m:sSub>
<m:r xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <m:t xml:space="preserve"> = </m:t>
</m:r>
<m:sSub xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <m:e><m:r><m:t>SNR</m:t></m:r></m:e>
  <m:sub><m:r><m:t>k</m:t></m:r></m:sub>
</m:sSub>
<m:r xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <m:t xml:space="preserve"> &#x22C5; </m:t>
</m:r>
<m:sSub xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <m:e><m:r><m:t>coh</m:t></m:r></m:e>
  <m:sub><m:r><m:t>k</m:t></m:r></m:sub>
</m:sSub>
<m:r xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <m:t xml:space="preserve"> &#x22C5; </m:t>
</m:r>
<m:sSub xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <m:e><m:r><m:t>ph_cons</m:t></m:r></m:e>
  <m:sub><m:r><m:t>k</m:t></m:r></m:sub>
</m:sSub>
"""

# ④ 状态量 [Δφ_k, φ̇_k]  (state vector for EKF)
OMML_STATE = """
<m:r xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <m:t xml:space="preserve">状态量：</m:t>
</m:r>
<m:d xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <m:dPr>
    <m:begChr m:val="["/>
    <m:endChr m:val="]"/>
    <m:ctrlPr/>
  </m:dPr>
  <m:e>
    <m:sSub>
      <m:e><m:r><m:rPr><m:sty m:val="i"/></m:rPr><m:t>&#x394;&#x3C6;</m:t></m:r></m:e>
      <m:sub><m:r><m:t>k</m:t></m:r></m:sub>
    </m:sSub>
    <m:r><m:t xml:space="preserve">, </m:t></m:r>
    <m:sSub>
      <m:e>
        <m:acc>
          <m:accPr>
            <m:chr m:val="&#x2D9;"/>
            <m:ctrlPr/>
          </m:accPr>
          <m:e><m:r><m:rPr><m:sty m:val="i"/></m:rPr><m:t>&#x3C6;</m:t></m:r></m:e>
        </m:acc>
      </m:e>
      <m:sub><m:r><m:t>k</m:t></m:r></m:sub>
    </m:sSub>
  </m:e>
</m:d>
"""


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


def hline(slide, x1, y, x2, color, width=Pt(0.6)):
    c = slide.shapes.add_connector(1, x1, y, x2, y)
    c.line.color.rgb = color
    c.line.width = width


def vline(slide, x, y1, y2, color, width=Pt(0.6)):
    c = slide.shapes.add_connector(1, x, y1, x, y2)
    c.line.color.rgb = color
    c.line.width = width


def tb(slide, text, l, t, w, h,
       sz=11, bold=False, color=DARK_TEXT,
       align=PP_ALIGN.LEFT, italic=False, wrap=True):
    box = slide.shapes.add_textbox(l, t, w, h)
    tf  = box.text_frame
    tf.word_wrap = wrap
    p   = tf.paragraphs[0]
    p.alignment = align
    r   = p.add_run()
    r.text = text
    r.font.size   = Pt(sz)
    r.font.bold   = bold
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
    p  = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r  = p.add_run()
    r.text = text
    r.font.size  = Pt(sz)
    r.font.bold  = bold
    r.font.color.rgb = tc


def add_omml_box(slide, omml_inner, l, t, w, h):
    """
    在 (l, t, w, h) 位置添加带有 OMML 方程的文本框，方程用淡蓝色背景衬托。
    omml_inner: m:oMath 内部 XML 片段（不含 m:oMath 标签）
    """
    # 淡蓝背景框（带蓝色边线）
    rect(slide, l, t, w, h, MATH_BG, lc=MATH_BD, lw=Pt(0.8))

    # 创建文本框
    box = slide.shapes.add_textbox(l + Inches(0.06), t + Inches(0.02),
                                    w - Inches(0.10), h - Inches(0.04))
    tf = box.text_frame
    tf.word_wrap = False

    # 获取段落 XML 元素
    p_elem = tf.paragraphs[0]._p

    # 清除已有内容（保留 pPr 属性）
    for child in list(p_elem):
        if child.tag != f'{{{NS_A}}}pPr':
            p_elem.remove(child)

    # 构建 a14:m 包装元素
    a14_m_xml = (
        '<a14:m'
        ' xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"'
        ' xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
        '<m:oMath>'
        + omml_inner +
        '</m:oMath>'
        '</a14:m>'
    )
    p_elem.append(etree.fromstring(a14_m_xml))
    return box


# ━━━━━━━━━━━━━━━━━━━━━━ 标题栏 & 页脚 ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def title_bar(slide, title_main, title_sub, page, total=1):
    rect(slide, Inches(0), Inches(0), Inches(0.18), TITLE_H, COL_NAVY)
    rect(slide, Inches(0.18), Inches(0), SW - Inches(0.18), TITLE_H, TITLE_BAR)
    hline(slide, Inches(0), TITLE_H, SW, TITLE_BAR2, Pt(2.5))
    tb(slide, title_main,
       Inches(0.36), Inches(0.10), Inches(10.5), Inches(0.38),
       sz=20, bold=True, color=WHITE)
    tb(slide, title_sub,
       Inches(0.36), Inches(0.50), Inches(10.5), Inches(0.26),
       sz=10.5, color=RGBColor(0xB0, 0xCC, 0xF0))
    circle_badge(slide, SW - Inches(0.52), TITLE_H / 2, 0.28,
                 TITLE_BAR2, f"{page}/{total}", sz=11)


def footer(slide, text):
    rect(slide, Inches(0), FOOTER_Y, SW, SH - FOOTER_Y, FOOTER_BG)
    hline(slide, MARGIN, FOOTER_Y, SW - MARGIN, TITLE_BAR2, Pt(1.5))
    rect(slide, MARGIN, FOOTER_Y + Inches(0.10),
         Inches(0.05), Inches(0.28), TITLE_BAR2)
    tb(slide, "  " + text,
       MARGIN + Inches(0.10), FOOTER_Y + Inches(0.09),
       CW - Inches(0.14), Inches(0.34),
       sz=8.8, italic=True, color=COL_NAVY)


# ━━━━━━━━━━━━━━━━━━━━━━ 子卡片（含 OMML 公式支持）━━━━━━━━━━━━━━━━━━

def sub_card(slide, l, t, w, h, accent, title, items,
             badge=None, title_sz=12, body_sz=10):
    """
    Sub-card with colored header band and white body.

    items: list of tuples with three possible formats:
        ("text", False)          — plain text bullet item
        ("text", True)           — text-only formula (navy color, indented)
        (omml_xml_str, "omml")   — OMML equation; omml_xml_str is the inner
                                   XML fragment (content of <m:oMath>) and
                                   renders as a proper MathType equation block
    """
    HDR_H     = Inches(0.36)
    MATH_H    = Inches(0.28)   # 公式行高
    BODY_LH   = body_sz * 1.80 * 914.4 / 72

    # 阴影
    rect(slide, l + Inches(0.03), t + Inches(0.03), w, h,
         RGBColor(0xC8, 0xD4, 0xEC))
    # 主体
    rect(slide, l, t, w, h, CARD_BG, lc=CARD_BORDER, lw=Pt(0.7))
    # 头带
    rect(slide, l, t, w, HDR_H, accent)

    if badge:
        circle_badge(slide, l + Inches(0.20), t + HDR_H / 2, 0.15,
                     WHITE, badge, sz=9, tc=accent)
    hx = l + (Inches(0.44) if badge else Inches(0.12))
    tb(slide, title, hx, t + Inches(0.04),
       w - hx + l - Inches(0.08), HDR_H - Inches(0.06),
       sz=title_sz, bold=True, color=WHITE)

    # 正文内容
    yy = t + HDR_H + Inches(0.06)
    PAD = Inches(0.10)

    for item in items:
        txt, kind = item[0], item[1]
        if kind == "omml":
            # 插入 OMML 方程块
            add_omml_box(slide, txt, l + PAD, yy,
                         w - 2 * PAD, MATH_H)
            yy += MATH_H + Inches(0.04)
        elif kind is True:
            # 旧式文本公式（navy 色等宽风格）
            tb(slide, "   " + txt,
               l + PAD, yy, w - 2 * PAD, BODY_LH,
               sz=body_sz - 0.5, color=COL_NAVY)
            yy += BODY_LH + Inches(0.01)
        else:
            # 普通文本
            tb(slide, "• " + txt,
               l + PAD, yy, w - 2 * PAD, BODY_LH,
               sz=body_sz, color=DARK_TEXT)
            yy += BODY_LH + Inches(0.01)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━ MAIN ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

prs = Presentation()
prs.slide_width  = SW
prs.slide_height = SH
blank = prs.slide_layouts[6]

s1 = prs.slides.add_slide(blank)
set_bg(s1, BG)

title_bar(s1,
          "多节点信号级高效相参合成 · 研究内容与技术路线",
          "Research Content & Technical Approach — MathType Equations Edition (v13)",
          1)

# ── 顶部流程带 ───────────────────────────────────────────────────────
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
       sz=9, bold=True, align=PP_ALIGN.CENTER, color=WHITE)
    if i < n_f - 1:
        tb(s1, "►", px + step_w, BOX_TOP + BOX_H * 0.25,
           ARR_W, BOX_H * 0.50,
           sz=10, color=COL_NAVY, align=PP_ALIGN.CENTER)
    px += step_w + ARR_W

# 指示竖线
for idx in [1, 2]:
    hx = MARGIN + idx * (step_w + ARR_W) + step_w / 2
    vline(s1, hx, BOX_TOP + BOX_H, BOX_TOP + BOX_H + Inches(0.14),
          TITLE_BAR2, Pt(1.2))

# ── 主内容区 ────────────────────────────────────────────────────────
MAIN_TOP = FLOW_TOP + FLOW_H + Inches(0.14)
MAIN_H   = CB - MAIN_TOP
GAP_MID  = Inches(0.16)

L_W = (CW - GAP_MID) * 0.48
R_X = MARGIN + L_W + GAP_MID
R_W = CW - L_W - GAP_MID

SEC_HDR_H = Inches(0.44)

# ── 左区大标题 ──
rect(s1, MARGIN + Inches(0.04), MAIN_TOP + Inches(0.04),
     L_W, SEC_HDR_H, RGBColor(0xBB, 0xCC, 0xBB))
rect(s1, MARGIN, MAIN_TOP, L_W, SEC_HDR_H, RC_HDR)
rect(s1, MARGIN, MAIN_TOP, Inches(0.10), SEC_HDR_H,
     RGBColor(0x80, 0xFF, 0xA0))
tb(s1, "📋",
   MARGIN + Inches(0.14), MAIN_TOP + Inches(0.06),
   Inches(0.30), SEC_HDR_H - Inches(0.08),
   sz=14, bold=True, color=WHITE)
tb(s1, "三项研究内容   Research Objectives",
   MARGIN + Inches(0.50), MAIN_TOP + Inches(0.06),
   L_W - Inches(0.56), Inches(0.26),
   sz=14, bold=True, color=WHITE)
tb(s1, "面向三个科学问题，分别设计可量化验证的研究任务",
   MARGIN + Inches(0.50), MAIN_TOP + Inches(0.28),
   L_W - Inches(0.56), Inches(0.14),
   sz=9, color=RGBColor(0xC0, 0xFF, 0xD0))

# 左区三个子卡片
SUB_GAP = Inches(0.08)
sub_top = MAIN_TOP + SEC_HDR_H + Inches(0.10)
sub_h   = (MAIN_H - SEC_HDR_H - Inches(0.10) - 2 * SUB_GAP) / 3

rc_cards = [
    {
        "accent": RC_SUB1,
        "badge":  "一",
        "title":  "增益退化端到端量化模型（对应科学问题 A）",
        "items": [
            ("建立时延误差 σ_τ / 相位误差 σ_φ / 频偏 σ_f → 合并增益 G 的统一传播链路", False),
            (OMML_GEFF, "omml"),   # ① G_eff 公式
            ('量化"效能90%保持区"：σ_φ < π/8，σ_τ < 0.5 ns，σ_f < 100 Hz', False),
            ("在 N=2,4,6 节点下仿真验证退化曲线，给出工程容差设计准则", False),
        ],
    },
    {
        "accent": RC_SUB2,
        "badge":  "二",
        "title":  "高动态时变相位误差的建模与实时估计",
        "items": [
            ("建立非平稳相位误差模型：", False),
            (OMML_PHASE, "omml"),  # ② Δφ_k 公式
            ("无导频自校准：基于回波互相关 + EKF 实时跟踪时变相位漂移率", False),
            ("推导最小可辨识帧长 N_c,min 与 SNR_k 的联合约束（CRLB 下界）", False),
        ],
    },
    {
        "accent": RC_SUB3,
        "badge":  "三",
        "title":  "质量自适应 MRC 合并与层次化处理框架",
        "items": [
            ("MRC 融合权值（三因子联合在线估计）：", False),
            (OMML_MRC, "omml"),    # ③ w_k 公式
            ("稳健机制：相位估计失败时自动降权，避免合并输出突变", False),
            ("分层策略：强目标 SIC 对消 → 弱目标三 MTI 检测 → 统一 MRC 合并", False),
        ],
    },
]

for ci, dc in enumerate(rc_cards):
    cy = sub_top + ci * (sub_h + SUB_GAP)
    sub_card(s1, MARGIN, cy, L_W, sub_h,
             dc["accent"], dc["title"], dc["items"],
             badge=dc["badge"])

# ── 右区大标题 ──
rect(s1, R_X + Inches(0.04), MAIN_TOP + Inches(0.04),
     R_W, SEC_HDR_H, RGBColor(0xC0, 0xC8, 0xE0))
rect(s1, R_X, MAIN_TOP, R_W, SEC_HDR_H, KT_HDR)
rect(s1, R_X, MAIN_TOP, Inches(0.10), SEC_HDR_H,
     RGBColor(0x80, 0xB8, 0xFF))
tb(s1, "⚙",
   R_X + Inches(0.14), MAIN_TOP + Inches(0.06),
   Inches(0.30), SEC_HDR_H - Inches(0.08),
   sz=14, bold=True, color=WHITE)
tb(s1, "关键技术方法与预期指标",
   R_X + Inches(0.50), MAIN_TOP + Inches(0.06),
   R_W - Inches(0.56), Inches(0.26),
   sz=14, bold=True, color=WHITE)
tb(s1, "Key Methods & Expected Technical Metrics — Quantifiable Validation Targets",
   R_X + Inches(0.50), MAIN_TOP + Inches(0.28),
   R_W - Inches(0.56), Inches(0.14),
   sz=9, color=RGBColor(0xB0, 0xD8, 0xFF))

# 右区三个子卡片
kt_cards = [
    {
        "accent": KT_SUB1,
        "badge":  "①",
        "title":  "误差传播链路建模与效能边界推导",
        "items": [
            ("理论工具：联合误差传播理论 + 贝叶斯效能下界（BCRLB）", False),
            ("输出：G_eff vs (σ_τ, σ_φ, σ_f) 三维效能图谱及工程容差曲线", False),
            ("预期：N=6 时理想 G=15.6 dB，误差容限内 G_eff ≥ 14 dB（90%）", False),
            ("验证：Monte-Carlo 仿真（1000次）+ MATLAB 端到端链路仿真", False),
        ],
    },
    {
        "accent": KT_SUB2,
        "badge":  "②",
        "title":  "EKF/自校准在线相位估计算法",
        "items": [
            (OMML_STATE, "omml"),   # ④ 状态向量
            ("量测更新：互相关亚像素峰值提供 Δτ 观测，EKF 步长 = 1 CPI", False),
            ("预期：SNR≥6 dB 时，相位估计误差 σ_est < π/8（G_eff ≥ 90%G）", False),
            ("对比：无补偿 / 静态互相关补偿 / 本方法三方案对比", False),
        ],
    },
    {
        "accent": KT_SUB3,
        "badge":  "③",
        "title":  "分层相参策略与实时计算优化",
        "items": [
            ("Round1：无 MTI 强目标检测（CFAR>12dB）→ SIC 对消 → 弱目标释放", False),
            ("Round2：三 MTI 并行（SLOW/FAST/VFAST）+ MRC → CFAR>6dB 统一输出", False),
            ("预期延迟：两轮合计 < 9 ms（Nc=64，Nsc=1024，N=6 链路）", False),
            ("优化：子带分解压缩 < 5 ms；精度损失 < 0.3 dB（仿真验证）", False),
        ],
    },
]

for ci, kc in enumerate(kt_cards):
    cy = sub_top + ci * (sub_h + SUB_GAP)
    sub_card(s1, R_X, cy, R_W, sub_h,
             kc["accent"], kc["title"], kc["items"],
             badge=kc["badge"])

footer(s1,
       "研究路径：量化退化律（内容一）→ 实时相位估计（内容二）→ 自适应 MRC 框架（内容三）"
       "——三项研究形成完整闭环，最终实现 N=6 节点下 G_eff ≥ 14 dB、延迟 < 9 ms 的工程目标")

# ── 保存 ───────────────────────────────────────────────────────────
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                   "coherent_synthesis_slides_v13.pptx")
prs.save(OUT)
print(f"✅  已生成：{OUT}")
