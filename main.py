import re
import io
import os
import uuid
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from fastapi import FastAPI, Request
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
from typing import Optional
from docx import Document
from docx.shared import Pt, Inches
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

# 环境字体设定
plt.rcParams['font.sans-serif'] = ['DejaVu Sans'] # 注意：若环境无中文字体，中文可能显示为框，生产环境建议安装对应字体
plt.rcParams['axes.unicode_minus'] = False

def set_font(run, size=12, bold=False):
    run.font.size = Pt(size)
    run.font.name = '宋体'
    run.bold = bold
    rPr = run._element.get_or_add_rPr()
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:eastAsia'), '宋体')
    rPr.append(rFonts)

def set_three_line_table(table):
    """学术三线表：顶线底线黑色 1.5pt，表头线 0.75pt"""
    thick_size = '12' 
    thin_size = '6'
    for i, row in enumerate(table.rows):
        for cell in row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            for border in ['top', 'bottom', 'left', 'right', 'insideH', 'insideV']:
                tag = f'w:{border}'
                element = tcPr.find(qn(tag))
                if element is not None: tcPr.remove(element)
            if i == 0:
                top = OxmlElement('w:top')
                top.set(qn('w:val'), 'single'); top.set(qn('w:sz'), thick_size); top.set(qn('w:color'), '000000')
                tcPr.append(top)
                bottom = OxmlElement('w:bottom')
                bottom.set(qn('w:val'), 'single'); bottom.set(qn('w:sz'), thin_size); bottom.set(qn('w:color'), '000000')
                tcPr.append(bottom)
            if i == len(table.rows) - 1:
                bottom = OxmlElement('w:bottom')
                bottom.set(qn('w:val'), 'single'); bottom.set(qn('w:sz'), thick_size); bottom.set(qn('w:color'), '000000')
                tcPr.append(bottom)

def draw_solid_pie(ax, values, labels):
    """绘制实心饼图：加粗大字体，无遮挡框"""
    # 颜色循环
    colors = plt.get_cmap('tab20c')(np.linspace(0, 1, len(values)))
    # 绘制实心饼图
    wedges, texts = ax.pie(values, startangle=-40, colors=colors)
    
    total = sum(values)
    for i, p in enumerate(wedges):
        ang = (p.theta2 - p.theta1)/2. + p.theta1
        y = np.sin(np.deg2rad(ang))
        x = np.cos(np.deg2rad(ang))
        
        # 决定文字方向
        horizontalalignment = {-1: "right", 1: "left"}[int(np.sign(x))]
        
        # 设置百分比和标签
        pct = values[i]/total*100
        val_str = f"{labels[i]}\n{int(values[i])} ({pct:.1f}%)"
        
        # 绘制指引线和文字（文字加粗 size=12）
        ax.annotate(val_str, xy=(x, y), xytext=(1.3*np.sign(x), 1.3*y),
                    horizontalalignment=horizontalalignment,
                    arrowprops=dict(arrowstyle="-", color="black", lw=0.8),
                    fontsize=12, fontweight='bold', color='black',
                    va="center")

def process_smart(doc, text):
    text = text.replace('$$', '').replace('\r', '')
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
                    fig, ax = plt.subplots(figsize=(10, 6)) # 略微调大画布
                    for col in df.columns[1:]:
                        df[col] = pd.to_numeric(df[col].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
                    
                    if ("分布" in current_title or "比例" in current_title) and fig_num not in [3, 6]:
                        draw_solid_pie(ax, df.iloc[:, 1].values, df.iloc[:, 0].values)
                        ax.set_xlim(-2.2, 2.2) # 扩大坐标轴范围防止文字截断
                        ax.set_ylim(-1.8, 1.8)
                    else:
                        x_indices = np.arange(len(df))
                        if len(df.columns) > 2:
                            width = 0.8 / (len(df.columns) - 1)
                            for i, col in enumerate(df.columns[1:]):
                                ax.bar(x_indices + i*width, df[col], width, label=col)
                            ax.set_xticks(x_indices + width*(len(df.columns)-2)/2)
                            ax.set_xticklabels(df.iloc[:, 0].astype(str), fontweight='bold')
                            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
                        else:
                            ax.bar(df.iloc[:, 0].astype(str), df.iloc[:, 1], color='#4472C4', width=0.5)
                            for i, val in enumerate(df.iloc[:, 1]):
                                ax.text(i, val, f'{int(val)}', ha='center', va='bottom', fontweight='bold', fontsize=11)
                    
                    plt.tight_layout(pad=3.0) # 增加内边距防止显示不全
                    buf = io.BytesIO(); plt.savefig(buf, format='png', dpi=180); plt.close(); buf.seek(0)
                    doc.add_picture(buf, width=Inches(5.8))
                    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
                else:
                    table = doc.add_table(rows=len(raw_data), cols=len(raw_data[0]))
                    table.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for i, row_data in enumerate(raw_data):
                        for j, val in enumerate(row_data):
                            cell = table.cell(i, j)
                            cell.text = re.sub(r'[^\w%\.\(\)]', '', str(val))
                            set_font(cell.paragraphs[0].runs[0], 10, i==0)
                    set_three_line_table(table)
            except: pass
        table_rows, is_chart_mode, current_title = [], False, ""

    for line in lines:
        l = line.strip()
        if not l: continue
        if l.startswith('#'):
            flush_table()
            p = doc.add_heading('', level=min(l.count('#'), 3))
            run = p.add_run(l.replace('#', '').strip())
            set_font(run, 14, True)
        elif re.match(r'(\*\*?)?[图表]\s?\d+[:：]', l):
            flush_table()
            current_title = l.replace('*', '').strip()
            is_chart_mode = "图" in current_title
            p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(current_title); set_font(run, 11, True)
        elif l.startswith('|'): table_rows.append(l)
        else:
            if table_rows: flush_table()
            p = doc.add_paragraph()
            p.paragraph_format.first_line_indent = Pt(24)
            run = p.add_run(l.replace('**', '')); set_font(run, 12)
    flush_table()

def add_page_number(doc):
    for section in doc.sections:
        footer = section.footer
        p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run()
        fldChar1 = OxmlElement('w:fldChar'); fldChar1.set(qn('w:fldCharType'), 'begin')
        instrText = OxmlElement('w:instrText'); instrText.set(qn('xml:space'), 'preserve'); instrText.text = 'PAGE'
        fldChar2 = OxmlElement('w:fldChar'); fldChar2.set(qn('w:fldCharType'), 'end')
        run._r.append(fldChar1); run._r.append(instrText); run._r.append(fldChar2)

def add_cover(doc):
    for _ in range(4): doc.add_paragraph()
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("山东师范大学人文社会科学科研成果发展态势分析报告\n（2020-2024）")
    set_font(run, size=22, bold=True); doc.add_page_break()

def add_toc(doc):
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("目录"); set_font(run, size=16, bold=True)
    p_toc = doc.add_paragraph(); run_toc = p_toc.add_run()
    fldChar1 = OxmlElement('w:fldChar'); fldChar1.set(qn('w:fldCharType'), 'begin')
    instrText = OxmlElement('w:instrText'); instrText.set(qn('xml:space'), 'preserve'); instrText.text = 'TOC \\o "1-3" \\h \\z \\u'
    fldChar2 = OxmlElement('w:fldChar'); fldChar2.set(qn('w:fldCharType'), 'separate')
    fldChar3 = OxmlElement('w:fldChar'); fldChar3.set(qn('w:fldCharType'), 'end')
    run_toc._r.append(fldChar1); run_toc._r.append(instrText); run_toc._r.append(fldChar2); run_toc._r.append(fldChar3)
    doc.add_page_break()

def force_update_fields(doc):
    element = doc.settings.element
    update_fields = OxmlElement('w:updateFields')
    update_fields.set(qn('w:val'), 'true')
    element.append(update_fields)

@app.post("/generate_report_word")
async def generate_report_word(input_data: ReportInput, request: Request):
    try:
        doc = Document()
        force_update_fields(doc); add_page_number(doc); add_cover(doc); add_toc(doc)
        has_content = False
        for i in range(1, 9): 
            txt = getattr(input_data, f"ch{i}_text", "")
            if txt and txt.strip():
                process_smart(doc, txt); has_content = True
        if not has_content: return {"file": "", "message": "输入为空", "status": "error"}
        fname = f"report_{uuid.uuid4().hex[:8]}.docx"
        full_path = os.path.join("static", fname)
        doc.save(full_path)
        file_url = f"{str(request.base_url).rstrip('/')}/static/{fname}"
        return {"file": file_url, "filename": fname, "status": "success"}
    except Exception as e: return {"file": "", "status": "error", "message": str(e)}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8080)
