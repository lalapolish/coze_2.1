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
from docx.shared import Pt, Inches, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_ALIGN_PARAGRAPH

app = FastAPI(openapi_version="3.0.0")

if not os.path.exists("static"):
    os.makedirs("static")
app.mount("/static", StaticFiles(directory="static"), name="static")

class ReportInput(BaseModel):
    ch1_text: Optional[str] = ""  
    ch2_text: Optional[str] = ""
    ch3_text: Optional[str] = ""
    ch4_text: Optional[str] = ""
    ch5_text: Optional[str] = ""
    ch6_text: Optional[str] = ""
    ch7_text: Optional[str] = ""
    ch8_text: Optional[str] = "" 
    appendix: Optional[str] = "" # 新增附录变量

# ================= 解决图片方框：强制加载本地 SimHei.ttf =================
font_path = os.path.join(os.getcwd(), 'SimHei.ttf')
if os.path.exists(font_path):
    fm.fontManager.addfont(font_path)
    plt.rcParams['font.sans-serif'] = ['SimHei']
else:
    plt.rcParams['font.sans-serif'] = ['SimHei', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

def set_font(run, size=12, bold=False):
    run.font.size = Pt(size)
    run.font.name = '宋体'
    run.bold = bold
    rPr = run._element.get_or_add_rPr()
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:eastAsia'), '宋体')
    rPr.append(rFonts)

def set_table_border(table):
    """为表格设置标准的黑色全边框"""
    tbl = table._tbl
    ptr = tbl.get_or_add_tblPr()
    borders = OxmlElement('w:tblBorders')
    for tag in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        edge = OxmlElement(f'w:{tag}')
        edge.set(qn('w:val'), 'single')
        edge.set(qn('w:sz'), '4') # 0.5pt
        edge.set(qn('w:color'), '000000')
        borders.append(edge)
    ptr.append(borders)

def draw_custom_pie(ax, values, labels):
    """饼图：彻底解决数值重叠问题"""
    colors = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47']
    wedges, _ = ax.pie(values, colors=colors[:len(values)], startangle=90, counterclock=False, 
                       wedgeprops={'edgecolor': 'white', 'linewidth': 1})
    
    total = sum(values)
    for i, p in enumerate(wedges):
        ang = (p.theta2 - p.theta1)/2. + p.theta1
        y = np.sin(np.deg2rad(ang))
        x = np.cos(np.deg2rad(ang))
        va = "center"; ha = "left" if x > 0 else "right"
        dist = 1.7
        y_text = 1.35 * y
        if abs(y) < 0.3: y_text = 1.5 * y
        connectionstyle = f"angle,angleA=0,angleB={ang}"
        label_text = f"{labels[i]}\n{int(values[i])} ({(values[i]/total*100):.1f}%)"
        ax.annotate(label_text, xy=(x, y), xytext=(dist*np.sign(x), y_text),
                    horizontalalignment=ha, verticalalignment=va,
                    arrowprops=dict(arrowstyle="-", color="black", connectionstyle=connectionstyle),
                    fontsize=10, fontweight='bold')

def add_page_number(doc):
    for sec in doc.sections:
        footer = sec.footer
        p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run()
        fldChar1 = OxmlElement('w:fldChar'); fldChar1.set(qn('w:fldCharType'), 'begin')
        instrText = OxmlElement('w:instrText'); instrText.set(qn('xml:space'), 'preserve'); instrText.text = 'PAGE'
        fldChar2 = OxmlElement('w:fldChar'); fldChar2.set(qn('w:fldCharType'), 'end')
        run._r.extend([fldChar1, instrText, fldChar2])

def process_smart(doc, text):
    text = text.replace('\\%', '%').replace('$$', '').replace('\r', '')
    text = re.sub(r'\([\d\.\+\-\*/\s]+\s*[≈=]\s*[\d\.\%]+\)', '', text)
    lines = text.split('\n')
    table_rows, current_title, is_chart_mode = [], "", False

    def flush_table():
        nonlocal table_rows, current_title, is_chart_mode
        if not table_rows: return
        raw_data = [[c.strip() for c in r.split('|') if c.strip()] for r in table_rows if '|' in r and '---' not in r]
        
        if len(raw_data) >= 2:
            try:
                df = pd.DataFrame(raw_data[1:], columns=raw_data[0])
                fig_num_match = re.search(r'\d+', current_title)
                fig_num = int(fig_num_match.group()) if fig_num_match else 0
                
                if is_chart_mode:
                    fig, ax = plt.subplots(figsize=(10, 6))
                    def clean(v):
                        s = re.sub(r'[^\d.]', '', str(v))
                        return float(s) if s else 0.0

                    # 1. 饼图 (4, 5)
                    if fig_num in [4, 5]:
                        v_list = [clean(v) for v in df.iloc[:, 1]]
                        draw_custom_pie(ax, v_list, df.iloc[:, 0].tolist())
                        ax.set_xlim(-2.5, 2.5); ax.set_ylim(-1.6, 1.6)

                    # 2. 多系列分组柱状图 (2, 12)
                    elif fig_num in [2, 12]:
                        df_plot = df[~df.iloc[:, 0].str.contains('合计|总计', na=False)]
                        x_labels = [re.sub(r'（.*?）|\(.*?\)|万元|项', '', str(x)) for x in df_plot.iloc[:, 0]]
                        categories = [c for c in df_plot.columns[1:] if '合计' not in c]
                        x = np.arange(len(x_labels))
                        width = 0.75 / len(categories)
                        for i, cat in enumerate(categories):
                            vals = [clean(v) for v in df_plot[cat]]
                            rects = ax.bar(x + i*width, vals, width, label=cat)
                            for rect in rects:
                                h = rect.get_height()
                                if h > 0: ax.text(rect.get_x()+rect.get_width()/2, h, f'{int(h)}', ha='center', va='bottom', fontsize=8)
                        ax.set_xticks(x + width*(len(categories)-1)/2); ax.set_xticklabels(x_labels)
                        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.12), ncol=len(categories))
                        plt.subplots_adjust(bottom=0.2)

                    # 3. 双轴图 (9)
                    elif fig_num == 9:
                        years = [re.sub(r'（.*?）|\(.*?\)', '', str(x)) for x in df.iloc[:, 0]]
                        counts = [clean(v) for v in df.iloc[:, 1]]; fundings = [clean(v) for v in df.iloc[:, 2]]
                        ax.bar(years, counts, color='#4472C4', label='立项数(项)', width=0.5)
                        ax2 = ax.twinx()
                        ax2.plot(years, fundings, color='#ED7D31', marker='o', linewidth=2, label='到账经费(万元)')
                        fig.legend(loc='lower center', bbox_to_anchor=(0.5, 0.02), ncol=2)
                        plt.subplots_adjust(top=0.88, bottom=0.15)

                    # 4. 横向趋势图 (10, 11)
                    elif fig_num in [10, 11]:
                        headers = [re.sub(r'[^\d]', '', str(h)) for h in df.columns]
                        year_headers = [h for h in headers if len(h) == 4]
                        if len(year_headers) >= 3:
                            x_labels = year_headers
                            y_vals = [clean(df.iloc[0, i]) for i in range(1, len(df.columns))]
                        else:
                            x_labels = [re.sub(r'（.*?）|\(.*?\)|万元', '', str(x)) for x in df.iloc[:, 0]]
                            y_vals = [clean(v) for v in df.iloc[:, 1]]
                        bars = ax.bar(x_labels, y_vals, color='#4472C4', width=0.5)
                        for bar in bars:
                            ax.text(bar.get_x()+bar.get_width()/2, bar.get_height(), f'{int(bar.get_height())}', ha='center', va='bottom', fontweight='bold')

                    # 5. 普通柱状图 (6, 8 等)
                    else:
                        x_labels = []
                        for x in df.iloc[:, 0]:
                            s = str(x).replace('万元', '').strip()
                            x_labels.append(s if s else "0")
                        
                        vals = [clean(v) for v in df.iloc[:, 1]]
                        ax.bar(x_labels, vals, color='#4472C4', width=0.5)
                        for i, v in enumerate(vals): ax.text(i, v, f'{int(v)}', ha='center', va='bottom')

                    buf = io.BytesIO(); plt.savefig(buf, format='png', dpi=150, bbox_inches='tight'); plt.close(); buf.seek(0)
                    doc.add_picture(buf, width=Inches(5.6))
                    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
                else:
                    # 表格逻辑
                    table = doc.add_table(rows=len(raw_data), cols=len(raw_data[0]))
                    table.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    table_num_match = re.search(r'\d+', current_title)
                    table_num = int(table_num_match.group()) if table_num_match else 0

                    for i, row_data in enumerate(raw_data):
                        row_str = "".join(row_data)
                        is_special = ("B级" in row_str or "C级" in row_str) and "发文数量" in row_str
                        
                        if is_special and table_num == 10:
                            merged_cell = table.cell(i, 0).merge(table.cell(i, len(row_data)-1))
                            for p in merged_cell.paragraphs: p.text = ""
                            p = merged_cell.paragraphs[0]
                            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            set_font(p.add_run(row_str.replace('*', '').strip()), 10, True)
                        else:
                            for j, val in enumerate(row_data):
                                cell = table.cell(i, j)
                                cell_text = val
                                if table_num == 6 and i == 0 and "总计" not in cell_text:
                                    cell_text = cell_text.replace("（万元）", "").replace("(万元)", "")
                                
                                cell.text = cell_text
                                p = cell.paragraphs[0]
                                set_font(p.runs[0] if p.runs else p.add_run(cell_text), 10, i==0)
                    set_table_border(table)
            except Exception as e: print(f"Error processing {current_title}: {e}")
        table_rows, is_chart_mode, current_title = [], False, ""

    for line in lines:
        l = line.strip()
        if not l: continue
        if l.startswith('#'):
            flush_table(); p = doc.add_heading('', level=min(l.count('#'), 3))
            run = p.add_run(l.replace('#', '').strip()); set_font(run, 14, True)
        elif re.match(r'(\*\*?)?[图表]\s?\d+[:：\s]', l):
            flush_table(); current_title = l.replace('*', '').strip(); is_chart_mode = "图" in current_title
            p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(current_title); set_font(run, 11, True)
        elif l.startswith('|'): table_rows.append(l)
        else:
            if table_rows: flush_table()
            p = doc.add_paragraph(); p.paragraph_format.first_line_indent = Pt(24)
            run = p.add_run(l.replace('**', '')); set_font(run, 12)
    flush_table()

def add_toc(doc):
    """插入目录域"""
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("目  录"); set_font(run, size=16, bold=True)
    
    p_toc = doc.add_paragraph()
    run_toc = p_toc.add_run()
    fldChar1 = OxmlElement('w:fldChar'); fldChar1.set(qn('w:fldCharType'), 'begin')
    instrText = OxmlElement('w:instrText'); instrText.set(qn('xml:space'), 'preserve')
    instrText.text = 'TOC \\o "1-3" \\h \\z \\u'
    fldChar2 = OxmlElement('w:fldChar'); fldChar2.set(qn('w:fldCharType'), 'separate')
    fldChar3 = OxmlElement('w:fldChar'); fldChar3.set(qn('w:fldCharType'), 'end')
    run_toc._r.extend([fldChar1, instrText, fldChar2, fldChar3])
    doc.add_page_break()

@app.post("/generate_report_word")
async def generate_report_word(input_data: ReportInput, request: Request):
    try:
        doc = Document()
        element = doc.settings.element
        update_fields = OxmlElement('w:updateFields'); update_fields.set(qn('w:val'), 'true'); element.append(update_fields)
        
        # 封面
        for _ in range(4): doc.add_paragraph()
        p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run("山东师范大学人文社会科学科研成果发展态势分析报告\n（2020-2024）")
        set_font(run, size=22, bold=True); doc.add_page_break()
        
        add_toc(doc)
        add_page_number(doc)
        
        # 处理正文 1-8 章
        for i in range(1, 9): 
            txt = getattr(input_data, f"ch{i}_text", "")
            if txt and txt.strip(): process_smart(doc, txt)
        
        # 处理附录部分
        if input_data.appendix and input_data.appendix.strip():
            doc.add_page_break() # 附录另起一页
            p_app = doc.add_paragraph()
            p_app.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run_app = p_app.add_run("附录")
            # 设置宋体、3号 (16pt)、加粗
            set_font(run_app, size=16, bold=True)
            # 使用 process_smart 解析附录中的 Markdown 表格和文字
            process_smart(doc, input_data.appendix)
        
        fname = f"report_{uuid.uuid4().hex[:8]}.docx"
        full_path = os.path.join("static", fname)
        doc.save(full_path)
        return {"file": f"{str(request.base_url).rstrip('/')}/static/{fname}", "status": "success"}
    except Exception as e: return {"file": "", "status": "error", "message": str(e)}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8080)
