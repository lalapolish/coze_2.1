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
from docx import Document
from docx.shared import Pt, Inches, RGBColor, Cm
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL

# ================= 1. 样式与基础工具 (100% 还原并植入新需求) =================
font_path = os.path.join(os.getcwd(), 'SimHei.ttf')
if os.path.exists(font_path):
    fm.fontManager.addfont(font_path)
    plt.rcParams['font.sans-serif'] = ['SimHei']
else:
    plt.rcParams['font.sans-serif'] = ['SimHei', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

def set_font(run, size=12, bold=False):
    """需求 (4): 强制黑色字体，支持中西文宋体"""
    run.font.size = Pt(size)
    run.font.name = '宋体'
    run.bold = bold
    run.font.color.rgb = RGBColor(0, 0, 0) # 强制全黑
    rPr = run._element.get_or_add_rPr()
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:eastAsia'), '宋体')
    rPr.append(rFonts)

def set_cell_shading(cell, color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color)
    tcPr.append(shd)

def set_table_border(table):
    tbl = table._tbl
    ptr = tbl.get_or_add_tblPr()
    borders = OxmlElement('w:tblBorders')
    for border_name in ['top', 'bottom', 'left', 'right']:
        edge = OxmlElement(f'w:{border_name}')
        edge.set(qn('w:val'), 'single')
        edge.set(qn('w:sz'), '12') 
        edge.set(qn('w:color'), '4472C4')
        borders.append(edge)
    inside_h = OxmlElement('w:insideH')
    inside_h.set(qn('w:val'), 'single')
    inside_h.set(qn('w:sz'), '4') 
    inside_h.set(qn('w:color'), '4472C4')
    borders.append(inside_h)
    ptr.append(borders)

# ================= 2. 绘图逻辑 (保留所有 fig_num 判定和复杂算法) =================
def draw_custom_pie(ax, values, labels, fig_num=0):
    colors = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47']
    wedges, _ = ax.pie(values, colors=colors[:len(values)], startangle=90, counterclock=False, 
                       wedgeprops={'edgecolor': 'white', 'linewidth': 1})
    total = sum(values)
    for i, p in enumerate(wedges):
        ang = (p.theta2 - p.theta1)/2. + p.theta1
        y, x = np.sin(np.deg2rad(ang)), np.cos(np.deg2rad(ang))
        dist, y_text = 1.7, 1.35 * y
        if abs(y) < 0.3: y_text = 1.5 * y
        x_text = dist * np.sign(x) if x != 0 else dist
        ha = "left" if x_text > 0 else "right"
        if fig_num == 5:
            lbl_str = str(labels[i]).strip()
            if lbl_str == 'A': x_text, ha, y_text = 1.7, "left", 1.6
            elif lbl_str == 'B': x_text, ha, y_text = -1.7, "right", 1.6
        label_text = f"{labels[i]}\n{int(values[i])} ({(values[i]/total*100):.1f}%)"
        ax.annotate(label_text, xy=(x, y), xytext=(x_text, y_text),
                    horizontalalignment=ha, verticalalignment="center",
                    arrowprops=dict(arrowstyle="-", color="black", connectionstyle=f"angle,angleA=0,angleB={ang}"),
                    fontsize=10, fontweight='bold', color='black')

# ================= 3. 核心解析与文档生成 (完整无删减) =================
def process_smart(doc, text):
    text = text.replace('\\%', '%').replace('$$', '').replace('\r', '')
    text = re.sub(r'\([\d\.\+\-\*/\s]+\s*[≈=]\s*[\d\.\%]+\)', '', text)
    lines = text.split('\n')
    table_rows, current_title, is_chart_mode = [], "", False

    def flush_table():
        nonlocal table_rows, current_title, is_chart_mode
        if not table_rows: return
        raw_data = []
        for r in table_rows:
            if '|' in r and '---' not in r:
                cells = [c.strip() for c in r.split('|') if c.strip()]
                if cells: raw_data.append(cells)
        
        if len(raw_data) >= 2:
            try:
                df = pd.DataFrame(raw_data[1:], columns=raw_data[0])
                fig_match = re.search(r'\d+', current_title)
                fig_num = int(fig_match.group()) if fig_match else 0
                
                if is_chart_mode:
                    fig, ax = plt.subplots(figsize=(10, 6))
                    # 需求 (3): 柱状图增加背景网格
                    if fig_num not in [4, 5]:
                        ax.yaxis.grid(True, linestyle='-', which='major', color='#D9D9D9', alpha=0.6)
                        ax.set_axisbelow(True)

                    def clean_val(v):
                        s = re.sub(r'[^\d.]', '', str(v))
                        return float(s) if s else 0.0

                    # --- 完整图表分支开始 ---
                    if fig_num in [4, 5]:
                        v_list = [clean_val(v) for v in df.iloc[:, 1]]
                        draw_custom_pie(ax, v_list, df.iloc[:, 0].tolist(), fig_num)
                        ax.set_xlim(-2.5, 2.5); ax.set_ylim(-1.6, 1.6)
                    elif fig_num in [2, 12]:
                        c_list = ['#F4CCCC', '#F9CB9C', '#FFE599', '#B6D7A8', '#A2C4C9', '#A4C2F4']
                        df_plot = df[~df.iloc[:, 0].str.contains('合计|总计', na=False)]
                        x_labs = [re.sub(r'（.*?）|\(.*?\)|万元|项', '', str(x)) for x in df_plot.iloc[:, 0]]
                        cats = [c for c in df_plot.columns[1:] if '合计' not in c]
                        x_idx = np.arange(len(x_labs))
                        w = 0.8 / (len(cats) + 1)
                        for idx, cat in enumerate(cats):
                            vals = [clean_val(v) for v in df_plot[cat]]
                            ax.bar(x_idx + idx*w, vals, w, label=cat, color=c_list[idx % len(c_list)])
                        ax.set_xticks(x_idx + w*(len(cats)-1)/2)
                        ax.set_xticklabels(x_labs, color='black')
                        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=len(cats))
                        plt.subplots_adjust(bottom=0.2)
                    elif fig_num == 9:
                        years = [re.sub(r'（.*?）|\(.*?\)', '', str(x)) for x in df.iloc[:, 0]]
                        cnts = [clean_val(v) for v in df.iloc[:, 1]]
                        funds = [clean_val(v) for v in df.iloc[:, 2]]
                        b_obj = ax.bar(years, cnts, color='#4472C4', label='立项数(项)', width=0.5)
                        for bar in b_obj:
                            ax.text(bar.get_x() + bar.get_width()/2, bar.get_height(), f'{int(bar.get_height())}', ha='center', va='bottom', color='black')
                        ax2 = ax.twinx()
                        ax2.plot(years, funds, color='#ED7D31', marker='o', linewidth=2, label='到账经费(万元)')
                        for i, v in enumerate(funds):
                            ax2.text(years[i], v, f'{v:.2f}', ha='center', va='bottom', color='black')
                        fig.legend(loc='lower center', bbox_to_anchor=(0.5, 0.02), ncol=2)
                        plt.subplots_adjust(top=0.88, bottom=0.15)
                    elif fig_num in [10, 11]:
                        heads = [re.sub(r'[^\d]', '', str(h)) for h in df.columns]
                        yr_heads = [h for h in heads if len(h) == 4]
                        if len(yr_heads) >= 3:
                            x_ls, y_vs = yr_heads, [clean_val(df.iloc[0, i]) for i in range(1, len(df.columns))]
                        else:
                            x_ls, y_vs = [re.sub(r'（.*?）|\(.*?\)', '', str(x)) for x in df.iloc[:, 0]], [clean_val(v) for v in df.iloc[:, 1]]
                        bars = ax.bar(x_ls, y_vs, color='#4472C4', width=0.5)
                        for b in bars:
                            ax.text(b.get_x()+b.get_width()/2, b.get_height(), f'{int(b.get_height())}', ha='center', va='bottom', color='black')
                    else:
                        x_ls = [str(x).replace('万元', '').strip() for x in df.iloc[:, 0]]
                        v_ls = [clean_val(v) for v in df.iloc[:, 1]]
                        bars = ax.bar(x_ls, v_ls, color='#4472C4', width=0.5)
                        for i, v in enumerate(v_ls):
                            ax.text(i, v, f'{int(v)}', ha='center', va='bottom', color='black')

                    lbl_map = {1:('发表年份','论文数量'), 3:('年份','立项数量'), 6:('经费区间（万元）','项目数量'), 7:('立项年份','项目数量'), 8:('经费区间','项目数量')}
                    if fig_num in lbl_map:
                        ax.set_xlabel(lbl_map[fig_num][0], fontweight='bold', color='black')
                        ax.set_ylabel(lbl_map[fig_num][1], fontweight='bold', color='black')

                    buf = io.BytesIO(); plt.savefig(buf, format='png', dpi=150, bbox_inches='tight'); plt.close(); buf.seek(0)
                    doc.add_picture(buf, width=Inches(5.6))
                    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
                else:
                    # --- 完整表格生成开始 ---
                    table = doc.add_table(rows=len(raw_data), cols=len(raw_data[0]))
                    table.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    table.allow_autofit = False
                    t_match = re.search(r'\d+', current_title)
                    t_num = int(t_match.group()) if t_match else 0

                    for i, r_data in enumerate(raw_data):
                        table.rows[i].height = Cm(0.71)
                        row_full_text = "".join(r_data)
                        # 需求 (2): 表 10 样式背景修正
                        is_rank_row = ("B级" in row_full_text or "C级" in row_full_text) and t_num == 10
                        if is_rank_row:
                            m_cell = table.cell(i, 0).merge(table.cell(i, len(r_data)-1))
                            for p in m_cell.paragraphs: p.clear()
                            p = m_cell.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            m_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                            set_font(p.add_run(row_full_text.replace('*', '').strip()), 11, True)
                            set_cell_shading(m_cell, "D9E1F2")
                        else:
                            for j, val in enumerate(r_data):
                                cell = table.cell(i, j)
                                # 完整列宽逻辑
                                head_c = str(raw_data[0][j])
                                if "序号" in head_c: cell.width = Cm(1.0)
                                elif "姓名" in head_c: cell.width = Cm(2.0)
                                elif any(x in head_c for x in ["单位", "学院"]): cell.width = Cm(4.8)
                                elif re.match(r'^20\d{2}$', val) or "年份" in head_c: cell.width = Cm(1.8)
                                else: cell.width = Cm(2.5)

                                cell.text = val; cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                                p = cell.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                                if i == 0:
                                    set_cell_shading(cell, "4472C4")
                                    set_font(p.runs[0] if p.runs else p.add_run(val), 11, True)
                                elif i % 2 == 0:
                                    set_cell_shading(cell, "D9E1F2")
                                    set_font(p.runs[0] if p.runs else p.add_run(val), 11, False)
                                else:
                                    set_font(p.runs[0] if p.runs else p.add_run(val), 11, False)
                    set_table_border(table)
            except Exception as e: print(f"Error in flush_table: {e}")
        table_rows, is_chart_mode, current_title = [], False, ""

    # --- 遍历逻辑 (带 1.5 倍行距需求) ---
    for line in lines:
        clean_l = line.strip()
        if not clean_l: continue
        if clean_l.startswith('#'):
            flush_table()
            h_count = clean_l.count('#')
            p = doc.add_heading('', level=min(h_count, 3)) if h_count <= 3 else doc.add_paragraph()
            p.paragraph_format.line_spacing = 1.5 # 需求 (1)
            if "附录" in clean_l: p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_font(p.add_run(clean_l.replace('#', '').strip()), 14 if h_count <= 3 else 12, True)
        elif re.match(r'(\*\*?)?(附)?[图表]\s?[\d\-\.]+[:：\s]', clean_l):
            flush_table(); current_title = clean_l.replace('*', '').strip(); is_chart_mode = "图" in current_title
            p = doc.add_paragraph(); p.paragraph_format.line_spacing = 1.5
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_font(p.add_run(current_title), 11, True)
        elif clean_l.startswith('|'): table_rows.append(clean_l)
        else:
            if table_rows: flush_table()
            p = doc.add_paragraph(); p.paragraph_format.line_spacing = 1.5 # 需求 (1)
            p.paragraph_format.first_line_indent = Pt(24)
            set_font(p.add_run(clean_l.replace('**', '')), 12)
    flush_table()

# ================= 4. 辅助函数 (目录/页码/页脚) =================
def add_toc(doc):
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.line_spacing = 1.5
    set_font(p.add_run("目  录"), size=16, bold=True)
    p_toc = doc.add_paragraph(); p_toc.paragraph_format.line_spacing = 1.5
    run_toc = p_toc.add_run()
    f1 = OxmlElement('w:fldChar'); f1.set(qn('w:fldCharType'), 'begin')
    inst = OxmlElement('w:instrText'); inst.set(qn('xml:space'), 'preserve'); inst.text = 'TOC \\o "1-3" \\h \\z \\u'
    f2 = OxmlElement('w:fldChar'); f2.set(qn('w:fldCharType'), 'separate')
    f3 = OxmlElement('w:fldChar'); f3.set(qn('w:fldCharType'), 'end')
    run_toc._r.extend([f1, inst, f2, f3]); doc.add_page_break()

def add_page_number(doc):
    for sec in doc.sections:
        footer = sec.footer
        p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run()
        f1 = OxmlElement('w:fldChar'); f1.set(qn('w:fldCharType'), 'begin')
        inst = OxmlElement('w:instrText'); inst.set(qn('xml:space'), 'preserve'); inst.text = 'PAGE'
        f2 = OxmlElement('w:fldChar'); f2.set(qn('w:fldCharType'), 'end')
        run._r.extend([f1, inst, f2])

# ================= 5. Workflow 入口函数 (解决链接问题) =================
def main(args):
    """
    args 是字典，包含 ch1_text, ch2_text... appendix 等上游变量
    """
    try:
        doc = Document()
        doc.settings.element.append(OxmlElement('w:updateFields')).set(qn('w:val'), 'true')
        
        # 封面 (保持 1.5 倍行距)
        for _ in range(4): 
            p = doc.add_paragraph(); p.paragraph_format.line_spacing = 1.5
        p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.line_spacing = 1.5
        set_font(p.add_run("山东师范大学人文社会科学科研成果发展态势分析报告\n（2020-2024）"), size=22, bold=True)
        doc.add_page_break()
        
        add_toc(doc)
        add_page_number(doc)
        
        # 遍历 1-8 章
        for i in range(1, 9):
            content = args.get(f"ch{i}_text", "")
            if content and str(content).strip():
                process_smart(doc, str(content))
        
        # 附录
        appendix = args.get("appendix", "")
        if appendix and str(appendix).strip():
            doc.add_page_break()
            process_smart(doc, str(appendix))
            
        # 生成并保存
        file_name = f"Report_{uuid.uuid4().hex[:8]}.docx"
        file_path = os.path.join(os.getcwd(), file_name)
        doc.save(file_path)
        
        # 直接返回文件路径，Workflow 会自动处理下载链接
        return {
            "file": file_path,
            "status": "success"
        }
    except Exception as e:
        return {"file": "", "status": "error", "log": str(e)}
