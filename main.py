import re
import io
import os
import uuid
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm

from fastapi import FastAPI, Request
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
from typing import Optional

from docx import Document
from docx.shared import Pt, Inches, RGBColor, Cm
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_ALIGN_VERTICAL, WD_ROW_HEIGHT_RULE


app = FastAPI()
app.openapi_version = "3.0.0"

if not os.path.exists("static"):
    os.makedirs("static")
app.mount("/static", StaticFiles(directory="static"), name="static")


# ================= 统一字号（五号 = 10.5pt） =================
TABLE_FONT_SIZE = 10.5        # 表格内所有文字
CAPTION_FONT_SIZE = 10.5      # 表标题 / 图标题
BAR_LABEL_FONT_SIZE = 15      # 柱状图柱顶数字
# ============================================================


class ReportInput(BaseModel):
    ch1_text: Optional[str] = ""
    ch2_text: Optional[str] = ""
    ch3_text: Optional[str] = ""
    ch4_text: Optional[str] = ""
    ch5_text: Optional[str] = ""
    ch6_text: Optional[str] = ""
    ch7_text: Optional[str] = ""
    ch8_text: Optional[str] = ""
    appendix: Optional[str] = ""


# ================= 解决图片方框：强制加载本地 SimHei.ttf =================
font_path = os.path.join(os.getcwd(), 'SimHei.ttf')
if os.path.exists(font_path):
    fm.fontManager.addfont(font_path)
    plt.rcParams['font.sans-serif'] = ['SimHei']
else:
    plt.rcParams['font.sans-serif'] = ['SimHei', 'DejaVu Sans']

plt.rcParams['axes.unicode_minus'] = False
plt.rcParams['font.size'] = 14
plt.rcParams['font.weight'] = 'bold'
plt.rcParams['axes.labelweight'] = 'bold'
plt.rcParams['axes.titleweight'] = 'bold'


def tight_text(text):
    """去掉中文与数字 / 字母之间多余的空格。

    仅处理“中文 ↔ ASCII”的交界处，英文单词之间、纯英文句子的空格不受影响。
    """
    if not text:
        return text
    s = str(text)
    # 中文（含全角标点）后面紧跟的数字 / 字母，去掉中间空格：“2816 篇” → “2816篇”
    s = re.sub(
        r'(?<=[\u3000-\u303f\u4e00-\u9fff，。、；：？！）】”’》])\s+(?=[0-9A-Za-z])',
        '', s
    )
    # 数字 / 字母后面紧跟中文，去掉中间空格：“F 级” → “F级”、“2024 年” → “2024年”
    s = re.sub(
        r'(?<=[0-9A-Za-z%\)\]】])\s+(?=[\u4e00-\u9fff，。、；：？！（【“‘《])',
        '', s
    )
    return s


def disable_auto_space_between_ascii_and_cjk(doc):
    """关闭 Word 默认在中文与西文 / 数字之间自动插入的间距。

    Word 的 autoSpaceDE / autoSpaceDN 会在“中文+数字”之间凭空撑出空隙，
    即使原文没有空格也会显示成“2816 篇”。在 docDefaults 里全局关掉。
    """
    try:
        styles_el = doc.styles.element
        doc_defaults = styles_el.find(qn('w:docDefaults'))
        if doc_defaults is None:
            return

        ppr_default = doc_defaults.find(qn('w:pPrDefault'))
        if ppr_default is None:
            ppr_default = OxmlElement('w:pPrDefault')
            doc_defaults.append(ppr_default)

        ppr = ppr_default.find(qn('w:pPr'))
        if ppr is None:
            ppr = OxmlElement('w:pPr')
            ppr_default.append(ppr)

        spacing = ppr.find(qn('w:spacing'))
        for tag in ('w:autoSpaceDE', 'w:autoSpaceDN'):
            el = ppr.find(qn(tag))
            if el is None:
                el = OxmlElement(tag)
                if spacing is not None:
                    spacing.addprevious(el)
                else:
                    ppr.append(el)
            el.set(qn('w:val'), '0')
    except Exception as e:
        print(f"Error disabling auto space: {e}")


def set_font(run, size=12, bold=False, color=None):
    """设置字体、大小、加粗及颜色"""
    run.font.size = Pt(size)
    run.font.name = '宋体'
    run.bold = bold
    if color:
        run.font.color.rgb = RGBColor.from_string(color)
    rPr = run._element.get_or_add_rPr()
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:eastAsia'), '宋体')
    rPr.append(rFonts)


def set_cell_shading(cell, color):
    """设置单元格背景颜色（同色重复设置时直接更新，避免重复写入 w:shd）"""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = tcPr.find(qn('w:shd'))
    if shd is None:
        shd = OxmlElement('w:shd')
        tcPr.append(shd)
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color)


def set_cell_bottom_border(cell, color="FFFFFF", size=12):
    """给单元格加下边框（用于表 10 表头与下一行之间加白色分隔线）"""
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = tcPr.find(qn('w:tcBorders'))
    if tcBorders is None:
        tcBorders = OxmlElement('w:tcBorders')
        # tcBorders 必须排在 w:shd 之前，否则 Word 可能忽略
        shd = tcPr.find(qn('w:shd'))
        if shd is not None:
            shd.addprevious(tcBorders)
        else:
            tcPr.append(tcBorders)

    bottom = tcBorders.find(qn('w:bottom'))
    if bottom is None:
        bottom = OxmlElement('w:bottom')
        tcBorders.append(bottom)
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), str(size))
    bottom.set(qn('w:space'), '0')
    bottom.set(qn('w:color'), color)


def set_table_border_no_vertical(table):
    """表格边框：保留横线，去掉竖线"""
    tbl = table._tbl
    tblPr = tbl.tblPr
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tbl.insert(0, tblPr)

    borders = OxmlElement('w:tblBorders')

    # 保留外边框上下线和内部横线，去掉竖线
    for border_name in ['top', 'bottom', 'insideH']:
        edge = OxmlElement(f'w:{border_name}')
        edge.set(qn('w:val'), 'single')
        edge.set(qn('w:sz'), '8')
        edge.set(qn('w:color'), '4472C4')
        borders.append(edge)

    # 竖线全部去掉
    for border_name in ['left', 'right', 'insideV']:
        edge = OxmlElement(f'w:{border_name}')
        edge.set(qn('w:val'), 'nil')
        borders.append(edge)

    tblPr.append(borders)


def set_table_fixed_layout(table):
    """设置表格固定布局，减少 Word 自动撑开/换行异常"""
    tblPr = table._tbl.tblPr
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        table._tbl.insert(0, tblPr)

    layout = OxmlElement('w:tblLayout')
    layout.set(qn('w:type'), 'fixed')
    tblPr.append(layout)


def set_table_row_height(row, min_height_cm=0.71):
    """
    自适应高度，但最小不低于 0.71cm
    这里使用 AT_LEAST，而不是 EXACTLY
    """
    row.height = Cm(min_height_cm)
    row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST


def set_paragraph_compact(paragraph, line_spacing=1.0, before=0, after=0):
    """设置段落紧凑，避免段间距过大"""
    paragraph.paragraph_format.space_before = Pt(before)
    paragraph.paragraph_format.space_after = Pt(after)
    paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    paragraph.paragraph_format.line_spacing = line_spacing


def set_paragraph_space_after_lines(paragraph, lines=0.5):
    """按“行”设置段后间距（Word 里显示为“段后 0.5 行”）。

    python-docx 的 space_after 只支持磅值（w:after），要精确表达“0.5 行”
    必须写 w:afterLines（单位为 1/100 行），同时把 w:after 归零。
    """
    pPr = paragraph._p.get_or_add_pPr()
    try:
        spacing = pPr.get_or_add_spacing()
    except Exception:
        spacing = pPr.find(qn('w:spacing'))
        if spacing is None:
            spacing = OxmlElement('w:spacing')
            pPr.append(spacing)
    spacing.set(qn('w:after'), '0')
    spacing.set(qn('w:afterLines'), str(int(round(lines * 100))))


def set_cell_text_compact(cell, font_size=10, bold=False, color="000000"):
    """统一设置单元格内文字紧凑样式"""
    for p in cell.paragraphs:
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
        p.paragraph_format.line_spacing = 1.0
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for r in p.runs:
            set_font(r, font_size, bold, color)


def style_data_source_paragraph(paragraph):
    """‘1.3 数据来源’下的正文段落：取消段间距"""
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    paragraph.paragraph_format.line_spacing = 1.0
    paragraph.paragraph_format.first_line_indent = Pt(24)


def apply_light_y_grid(fig, ax, grid_color="#D9D9D9", grid_width=1.0):
    """给柱状图加浅色横向网格线（条纹效果），并隐藏刻度短线

    - 只画 y 轴方向的网格线，x 轴保持干净
    - set_axisbelow(True) 保证网格线在柱子下方，不会压住柱子
    - 双轴图（twinx）的副轴网格全部关闭，避免出现重复线条
    """
    for a in fig.axes:
        a.grid(False)
        a.tick_params(axis='both', which='both', length=0)

    ax.set_axisbelow(True)
    ax.grid(axis='y', color=grid_color, linestyle='-', linewidth=grid_width, alpha=1.0)


def draw_custom_pie(ax, values, labels, fig_num=0):
    """饼图：短直线且不拐弯"""
    colors = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47']
    wedges, _ = ax.pie(
        values,
        colors=colors[:len(values)],
        startangle=90,
        counterclock=False,
        wedgeprops={'edgecolor': 'white', 'linewidth': 1}
    )

    total = sum(values)

    for i, p in enumerate(wedges):
        ang = (p.theta2 - p.theta1) / 2.0 + p.theta1
        y = np.sin(np.deg2rad(ang))
        x = np.cos(np.deg2rad(ang))

        dist = 1.25
        y_text = dist * y
        x_text = dist * x

        label = str(labels[i]).strip()
        display_label = label if label.endswith('级') else f"{label}级"

        if fig_num == 5:
            if label.startswith('A'):
                ha = "left"
                x_text = max(x_text, 0.10) + 0.18
                y_text = y_text + 0.08
            elif label.startswith('B'):
                ha = "right"
                x_text = min(x_text, -0.10) - 0.18
                y_text = y_text + 0.02
            elif label.startswith('D'):
                ha = "right"
                x_text = x_text - 0.20
                y_text = y
            else:
                ha = "left" if x > 0 else "right"
        elif fig_num == 4:
            ha = "left" if x > 0 else "right"
            if x > 0:
                x_text = x_text + 0.10
            else:
                x_text = x_text - 0.10
            y_text = y_text + 0.02
        else:
            ha = "left" if x > 0 else "right"

        connectionstyle = "arc3,rad=0"
        label_text = f"{display_label}\n{int(values[i])} ({(values[i] / total * 100):.1f}%)"

        ax.annotate(
            label_text,
            xy=(x, y),
            xytext=(x_text, y_text),
            horizontalalignment=ha,
            verticalalignment="center",
            rotation=0,
            rotation_mode='anchor',
            arrowprops=dict(
                arrowstyle="-",
                color="black",
                connectionstyle=connectionstyle
            ),
            fontsize=14,
            fontweight='bold'
        )


def add_page_number(doc):
    for sec in doc.sections:
        footer = sec.footer
        p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run()
        fldChar1 = OxmlElement('w:fldChar')
        fldChar1.set(qn('w:fldCharType'), 'begin')
        instrText = OxmlElement('w:instrText')
        instrText.set(qn('xml:space'), 'preserve')
        instrText.text = 'PAGE'
        fldChar2 = OxmlElement('w:fldChar')
        fldChar2.set(qn('w:fldCharType'), 'end')
        run._r.extend([fldChar1, instrText, fldChar2])


def process_smart(doc, text):
    text = text.replace('\\%', '%').replace('$$', '').replace('\r', '')
    text = re.sub(r'\([\d\.\+\-\*/\s]+\s*[≈=]\s*[\d\.\%]+\)', '', text)
    lines = text.split('\n')
    table_rows, current_title, is_chart_mode = [], "", False
    in_data_source = False

    def flush_table():
        nonlocal table_rows, current_title, is_chart_mode
        if not table_rows:
            return

        raw_data = []
        for r in table_rows:
            if '|' in r and '---' not in r:
                cells = [c.strip() for c in r.split('|')]
                if len(cells) >= 3:
                    raw_data.append(cells[1:-1])

        if len(raw_data) >= 2:
            try:
                num_match = re.search(r'\d+', current_title)
                current_num = int(num_match.group()) if num_match else 0

                if is_chart_mode:
                    df = pd.DataFrame(raw_data[1:], columns=raw_data[0])
                    fig, ax = plt.subplots(figsize=(10, 6))

                    for spine in ax.spines.values():
                        spine.set_visible(False)

                    def clean(v):
                        s = re.sub(r'[^\d.]', '', str(v))
                        return float(s) if s else 0.0

                    if current_num in [4, 5]:
                        v_list = [clean(v) for v in df.iloc[:, 1]]
                        draw_custom_pie(ax, v_list, df.iloc[:, 0].tolist(), current_num)
                        ax.set_xlim(-2.0, 2.0)
                        ax.set_ylim(-1.5, 1.5)

                    elif current_num in [2, 12]:
                        level_colors = {
                            'A': (229/255, 76/255, 94/255),
                            'B': (72/255, 116/255, 203/255),
                            'C': (238/255, 130/255, 47/255),
                            'D': (242/255, 186/255, 2/255),
                            'E': (117/255, 189/255, 66/255),
                            'F': (48/255, 192/255, 180/255)
                        }
                        df_plot = df[~df.iloc[:, 0].astype(str).str.contains('合计|总计', na=False)]
                        x_labels = [re.sub(r'（.*?）|\(.*?\)|万元|项', '', str(x)) for x in df_plot.iloc[:, 0]]
                        categories = [c for c in df_plot.columns[1:] if '合计' not in c and str(c).strip()]
                        x = np.arange(len(x_labels))
                        width = 0.6 / max(len(categories), 1)

                        for i, cat in enumerate(categories):
                            vals = [clean(v) for v in df_plot[cat]]
                            c_key = str(cat).strip().upper()[0]
                            bar_color = level_colors.get(c_key, (72/255, 116/255, 203/255))
                            rects = ax.bar(x + i * width, vals, width, label=cat, color=bar_color)
                            ax.bar_label(rects, padding=3, fontsize=BAR_LABEL_FONT_SIZE, fontweight='bold')

                        ax.set_xticks(x + width * (len(categories) - 1) / 2)
                        ax.set_xticklabels(x_labels, fontsize=13, fontweight='bold')
                        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=len(categories), frameon=False)
                        plt.subplots_adjust(bottom=0.2)

                    elif current_num == 9:
                        years = [re.sub(r'（.*?）|\(.*?\)', '', str(x)) for x in df.iloc[:, 0]]
                        counts = [clean(v) for v in df.iloc[:, 1]]
                        fundings = [clean(v) for v in df.iloc[:, 2]]

                        bars = ax.bar(years, counts, color='#4472C4', label='立项数(项)', width=0.38)
                        ax.bar_label(bars, padding=3, fontsize=BAR_LABEL_FONT_SIZE, fontweight='bold')

                        ax2 = ax.twinx()
                        for spine in ax2.spines.values():
                            spine.set_visible(False)

                        ax2.plot(years, fundings, color='#ED7D31', marker='o', linewidth=2, label='到账经费(万元)')

                        for i, val in enumerate(fundings):
                            ax2.annotate(
                                f'{val:.2f}',
                                xy=(years[i], val),
                                xytext=(14, -12),
                                textcoords='offset points',
                                ha='left',
                                va='top',
                                fontsize=12,
                                fontweight='bold',
                                color='red',
                                arrowprops=dict(
                                    arrowstyle='-',
                                    color='red',
                                    lw=1.5,
                                    shrinkA=0,
                                    shrinkB=0,
                                    connectionstyle='arc3,rad=0'
                                )
                            )

                        fig.legend(loc='lower center', bbox_to_anchor=(0.5, 0.00), ncol=2, frameon=False)
                        plt.subplots_adjust(top=0.88, bottom=0.20)

                    elif current_num in [10, 11]:
                        headers = [re.sub(r'[^\d]', '', str(h)) for h in df.columns]
                        year_headers = [h for h in headers if len(h) == 4]
                        if len(year_headers) >= 3:
                            x_labels = year_headers
                            y_vals = [clean(df.iloc[0, i]) for i in range(1, len(df.columns))]
                        else:
                            x_labels = [re.sub(r'（.*?）|\(.*?\)|万元', '', str(x)) for x in df.iloc[:, 0]]
                            y_vals = [clean(v) for v in df.iloc[:, 1]]

                        bars = ax.bar(x_labels, y_vals, color='#4472C4', width=0.38)
                        ax.bar_label(bars, padding=3, fontsize=BAR_LABEL_FONT_SIZE, fontweight='bold')

                    else:
                        x_labels = [str(x).replace('万元', '').strip() for x in df.iloc[:, 0]]
                        vals = [clean(v) for v in df.iloc[:, 1]]
                        bars = ax.bar(x_labels, vals, color='#4472C4', width=0.38)
                        ax.bar_label(bars, padding=3, fontsize=BAR_LABEL_FONT_SIZE, fontweight='bold')

                    xy_labels = {
                        1: ('发表年份', '论文数量'),
                        2: ('年份', '论文数量'),
                        3: ('年份', '立项数量'),
                        6: ('经费区间（万元）', '项目数量'),
                        7: ('立项年份', '项目数量'),
                        8: ('经费区间', '项目数量'),
                        9: ('年份', '项目数量'),
                        10: ('年份', '著作出版数量'),
                        11: ('年份', '获奖数量'),
                        12: ('年份', '获奖数量')
                    }
                    if current_num in xy_labels:
                        ax.set_xlabel(xy_labels[current_num][0], fontweight='bold', fontsize=17, labelpad=8)
                        ax.set_ylabel(xy_labels[current_num][1], fontweight='bold', fontsize=17, labelpad=8)
                        ax.tick_params(axis='both', labelsize=13)
                        for tick in ax.get_xticklabels():
                            tick.set_fontweight('bold')
                        for tick in ax.get_yticklabels():
                            tick.set_fontweight('bold')

                    # ================= 新增：柱状图加浅色横向网格线（条纹） =================
                    # 饼图（4、5）不加网格，其余柱状图/柱线组合图统一加
                    if current_num not in [4, 5]:
                        apply_light_y_grid(fig, ax, grid_color="#D9D9D9", grid_width=1.0)
                    # ======================================================================

                    buf = io.BytesIO()
                    plt.savefig(buf, format='png', dpi=150, bbox_inches='tight')
                    plt.close()
                    buf.seek(0)
                    doc.add_picture(buf, width=Inches(5.6))
                    # 图片所在段落设为紧凑，消除图片与图标题之间的空隙
                    pic_para = doc.paragraphs[-1]
                    set_paragraph_compact(pic_para, line_spacing=1.0, before=0, after=0)
                    pic_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

                    p = doc.add_paragraph()
                    set_paragraph_compact(p, line_spacing=1.0, before=0, after=0)
                    # 图标题：段后 0.5 行（段前保持 0）
                    set_paragraph_space_after_lines(p, 0.5)
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    run = p.add_run(current_title)
                    set_font(run, CAPTION_FONT_SIZE, True, "000000")

                else:
                    if current_num in [11, 13, 14]:
                        if len(raw_data) > 11:
                            header = raw_data[0]
                            data_body = raw_data[1:]
                            n = len(data_body)
                            half = (n + 1) // 2
                            new_raw_data = [header + header]
                            for i in range(half):
                                row_left = data_body[i]
                                row_right = data_body[i + half] if (i + half) < n else [""] * len(header)
                                new_raw_data.append(row_left + row_right)
                            raw_data = new_raw_data

                    if current_num == 12:
                        df_12 = pd.DataFrame(raw_data[1:], columns=raw_data[0])
                        new_table_data = [["学者", "所属单位", "项目数量", "学者", "所属单位", "项目数量"]]
                        levels = ["国家级", "省部级"]
                        for lvl in levels:
                            sub_df = df_12[df_12['级别'].astype(str).str.contains(lvl, na=False)]
                            if not sub_df.empty:
                                suffix = "（大于等于2项）" if lvl == "国家级" else "（大于等于3项）"
                                new_table_data.append([lvl + suffix] * 6)
                                sub_rows = sub_df[['负责人', '单位', '立项数']].values.tolist()
                                half = (len(sub_rows) + 1) // 2
                                for i in range(half):
                                    left = sub_rows[i]
                                    right = sub_rows[i + half] if (i + half) < len(sub_rows) else ["", "", ""]
                                    new_table_data.append(left + right)
                        raw_data = new_table_data

                    # ===== 表 9：两层表头（“认定等级”横跨 B~总计 六列）=====
                    is_two_level_9 = (
                        current_num == 9
                        and not is_chart_mode
                        and len(raw_data) > 1
                        and len(raw_data[0]) >= 9
                    )
                    if is_two_level_9:
                        _header9 = raw_data[0]
                        raw_data = [_header9, [""] * len(_header9)] + raw_data[1:]
                    # ========================================================

                    table = doc.add_table(rows=len(raw_data), cols=len(raw_data[0]))
                    table.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    table.autofit = True
                    table.allow_autofit = True
                    set_table_fixed_layout(table)

                    # 改为自适应高度，但最小高度不低于 0.71cm
                    for row in table.rows:
                        set_table_row_height(row, 0.71)

                    for i, row_values in enumerate(raw_data):
                        if is_two_level_9 and i == 1:
                            # 第二层表头留空，稍后与第一层合并后单独写入
                            continue

                        row_str = "".join([str(x) for x in row_values])
                        is_special_10 = ("B级" in row_str or "C级" in row_str) and current_num == 10
                        is_special_12 = ("国家级" in row_str or "省部级" in row_str) and current_num == 12 and len(set(row_values)) == 1

                        if is_special_10 or is_special_12:
                            display_text = tight_text(
                                next((x for x in row_values if str(x).strip()), "").replace('*', '').strip()
                            )
                            merged_cell = table.cell(i, 0).merge(table.cell(i, len(row_values) - 1))
                            merged_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                            set_cell_shading(merged_cell, "4472C4")

                            merged_cell.text = ""
                            p = merged_cell.paragraphs[0]
                            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            p.paragraph_format.space_before = Pt(0)
                            p.paragraph_format.space_after = Pt(0)
                            p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
                            p.paragraph_format.line_spacing = 1.0

                            run = p.add_run(display_text)
                            set_font(run, TABLE_FONT_SIZE, True, "FFFFFF")
                        else:
                            for j, val in enumerate(row_values):
                                if j >= len(table.columns):
                                    break

                                cell = table.cell(i, j)
                                cell.text = tight_text(str(val).replace('*', ''))
                                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

                                p = cell.paragraphs[0]
                                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                                p.paragraph_format.space_before = Pt(0)
                                p.paragraph_format.space_after = Pt(0)
                                p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
                                p.paragraph_format.line_spacing = 1.0

                                if i == 0:
                                    set_cell_shading(cell, "4472C4")
                                    if p.runs:
                                        set_font(p.runs[0], TABLE_FONT_SIZE, True, "FFFFFF")
                                    else:
                                        set_font(p.add_run(str(val)), TABLE_FONT_SIZE, True, "FFFFFF")
                                    # 表 10：表头底部加白色横线，与下方分类行区分
                                    if current_num == 10:
                                        set_cell_bottom_border(cell, "FFFFFF", 12)
                                elif i % 2 == 0:
                                    set_cell_shading(cell, "D9E1F2")
                                    if p.runs:
                                        set_font(p.runs[0], TABLE_FONT_SIZE, False, "000000")
                                    else:
                                        set_font(p.add_run(str(val)), TABLE_FONT_SIZE, False, "000000")
                                else:
                                    if p.runs:
                                        set_font(p.runs[0], TABLE_FONT_SIZE, False, "000000")
                                    else:
                                        set_font(p.add_run(str(val)), TABLE_FONT_SIZE, False, "000000")

                    # ===== 表 9 两层表头合并 =====
                    if is_two_level_9:
                        ncol = len(raw_data[0])
                        header_row = raw_data[0]

                        # 第一层：第 3 列起横向合并成“认定等级”
                        span = table.cell(0, 3).merge(table.cell(0, ncol - 1))
                        span.text = ""
                        p = span.paragraphs[0]
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        p.paragraph_format.space_before = Pt(0)
                        p.paragraph_format.space_after = Pt(0)
                        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
                        p.paragraph_format.line_spacing = 1.0
                        set_cell_shading(span, "4472C4")
                        span.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                        run = p.add_run("认定等级")
                        set_font(run, TABLE_FONT_SIZE, True, "FFFFFF")

                        # 前 3 列（序号 / 姓名 / 所属单位）纵向合并两行
                        for c in range(3):
                            merged = table.cell(0, c).merge(table.cell(1, c))
                            merged.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                            # 合并后 row1 的空段落会撑高行，删掉多余段落
                            for extra in merged.paragraphs[1:]:
                                extra._element.getparent().remove(extra._element)
                            set_cell_shading(merged, "4472C4")

                        # 第二层：B级 ~ 总计
                        for j in range(3, ncol):
                            cell = table.cell(1, j)
                            cell.text = ""
                            p = cell.paragraphs[0]
                            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            p.paragraph_format.space_before = Pt(0)
                            p.paragraph_format.space_after = Pt(0)
                            p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
                            p.paragraph_format.line_spacing = 1.0
                            set_cell_shading(cell, "4472C4")
                            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                            run = p.add_run(str(header_row[j]))
                            set_font(run, TABLE_FONT_SIZE, True, "FFFFFF")
                    # ==============================

                    set_table_border_no_vertical(table)

            except Exception as e:
                print(f"Error processing {current_title}: {e}")

        table_rows, is_chart_mode, current_title = [], False, ""

    for line in lines:
        l = line.strip().replace('<center>', '').replace('</center>', '')
        if not l:
            continue

        # 识别“1.3 数据来源”
        if re.match(r'^\d+\.\d+\s*数据来源', l):
            flush_table()
            in_data_source = True
            p = doc.add_paragraph()
            set_paragraph_compact(p, line_spacing=1.0, before=0, after=0)
            run = p.add_run(l)
            set_font(run, 14, True, "000000")
            continue

        if l.startswith('#'):
            flush_table()
            in_data_source = False
            hash_count = l.count('#')
            if hash_count <= 3:
                p = doc.add_heading('', level=hash_count)
                p.paragraph_format.line_spacing = 1.5
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
                if "附录" in l:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = p.add_run(l.replace('#', '').strip())
                set_font(run, 14, True, "000000")
            else:
                p = doc.add_paragraph()
                set_paragraph_compact(p, line_spacing=1.0, before=0, after=0)
                run = p.add_run(l.replace('#', '').strip())
                set_font(run, 12, True, "000000")

        elif re.match(r'(\*\*?)?(附)?[图表]\s?[\d\-\.]+[:：\s]', l):
            flush_table()
            current_title = l.replace('*', '').strip()
            is_chart_mode = "图" in current_title
            in_data_source = False
            if not is_chart_mode:
                p = doc.add_paragraph()
                set_paragraph_compact(p, line_spacing=1.0, before=0, after=0)
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = p.add_run(current_title)
                # 表标题：五号加粗
                set_font(run, CAPTION_FONT_SIZE, True, "000000")

        elif l.startswith('|'):
            table_rows.append(l)

        else:
            if table_rows:
                flush_table()

            p = doc.add_paragraph()
            if in_data_source:
                style_data_source_paragraph(p)
            else:
                p.paragraph_format.line_spacing = 1.5
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
                p.paragraph_format.first_line_indent = Pt(24)

            # “（1）”“（2）”这类小标题加粗
            is_sub_heading = bool(
                re.match(r'^[（(]\s*[0-9一二三四五六七八九十]+\s*[)）]', l)
            ) or l.startswith('**')

            run = p.add_run(tight_text(l.replace('**', '')))
            set_font(run, 12, is_sub_heading, "000000")

    flush_table()


def add_toc(doc):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("目  录")
    set_font(run, size=16, bold=True, color="000000")

    p_toc = doc.add_paragraph()
    run_toc = p_toc.add_run()

    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')

    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = 'TOC \\o "1-3" \\h \\z \\u'

    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'separate')

    hint_text = OxmlElement('w:t')
    hint_text.text = "（请在此处右键单击，选择“更新域”以生成目录内容）"

    fldChar3 = OxmlElement('w:fldChar')
    fldChar3.set(qn('w:fldCharType'), 'end')

    run_toc._r.extend([fldChar1, instrText, fldChar2, hint_text, fldChar3])
    doc.add_page_break()


@app.post("/generate_report_word")
async def generate_report_word(input_data: ReportInput, request: Request):
    try:
        doc = Document()

        # 关闭 Word 中文与数字/字母之间的自动间距
        disable_auto_space_between_ascii_and_cjk(doc)

        element = doc.settings.element
        update_fields = OxmlElement('w:updateFields')
        update_fields.set(qn('w:val'), 'true')
        element.append(update_fields)

        for _ in range(4):
            doc.add_paragraph()

        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run("山东师范大学人文社会科学科研成果发展态势分析报告\n（2020-2024）")
        set_font(run, size=22, bold=True, color="000000")

        doc.add_page_break()
        add_toc(doc)
        add_page_number(doc)

        for i in range(1, 9):
            txt = getattr(input_data, f"ch{i}_text", "")
            if txt and txt.strip():
                process_smart(doc, txt)

        if input_data.appendix and input_data.appendix.strip():
            doc.add_page_break()
            process_smart(doc, input_data.appendix)

        fname = f"report_{uuid.uuid4().hex[:8]}.docx"
        full_path = os.path.join("static", fname)
        doc.save(full_path)

        return {"file": f"{str(request.base_url).rstrip('/')}/static/{fname}", "status": "success"}

    except Exception as e:
        return {"file": "", "status": "error", "message": str(e)}


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8080)
