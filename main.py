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

# 设置 Matplotlib 字体
plt.rcParams['font.sans-serif'] = ['DejaVu Sans']
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
    """设置标准的黑色加粗三线表"""
    thick_sz = '18' # 2.25pt
    thin_sz = '8'   # 1pt
    color_val = '000000'

    for i, row in enumerate(table.rows):
        for cell in row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            for b in ['top', 'bottom', 'left', 'right', 'insideH', 'insideV']:
                tag = f'w:{b}'
                element = tcPr.find(qn(tag))
                if element is not None: tcPr.remove(element)
            
            if i == 0:
                top = OxmlElement('w:top')
                top.set(qn('w:val'), 'single'); top.set(qn('w:sz'), thick_sz); top.set(qn('w:color'), color_val)
                tcPr.append(top)
                btm = OxmlElement('w:bottom')
                btm.set(qn('w:val'), 'single'); btm.set(qn('w:sz'), thin_sz); btm.set(qn('w:color'), color_val)
                tcPr.append(btm)
            if i == len(table.rows) - 1:
                btm = OxmlElement('w:bottom')
                btm.set(qn('w:val'), 'single'); btm.set(qn('w:sz'), thick_sz); btm.set(qn('w:color'), color_val)
                tcPr.append(btm)

def draw_custom_pie(ax, values, labels):
    """饼图优化：防遮挡"""
    custom_colors = ['#4472C4', '#ED7D31', '#FFD966', '#E8307E', '#2EBB9F', '#2E3192']
    wedges, _ = ax.pie(values, colors=custom_colors[:len(values)], startangle=90, 
                       counterclock=False, wedgeprops={'edgecolor': 'white', 'linewidth': 1})
    
    total = sum(values)
    for i, p in enumerate(wedges):
        ang = (p.theta2 - p.theta1)/2. + p.theta1
        y = np.sin(np.deg2rad(ang))
        x = np.cos(np.deg2rad(ang))
        label_text = f"{labels[i]}\n{int(values[i])} ({(values[i]/total*100):.1f}%)"
        ax.annotate(label_text, xy=(x, y), xytext=(1.4*np.sign(x), 1.4*y),
                    horizontalalignment='center' if abs(x) < 0.1 else ('left' if x > 0 else 'right'),
                    arrowprops=dict(arrowstyle="-", color="black"),
                    fontsize=11, fontweight='bold', va='center')

def process_smart(doc, text):
    # 1. 强力清洗：去掉 LaTeX 转义、数学符号
    text = text.replace('\\%', '%').replace('$$', '').replace('\\approx', '≈').replace('\\times', '×')
    # 2. 正则清洗：去掉 AI 喜欢附带的计算公式，如 (12+34)/100 ≈ 46%
    text = re.sub(r'\([\d\.\+\-\*/\s]+\s*[≈=]\s*[\d\.\%]+\)', '', text)
    # 3. 清理空行多余空格
    text = text.replace('\r', '')
    
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
                # 预处理：剔除“总计”或“合计”列，防止绘图失衡
                cols_to_keep = [c for c in df.columns if "总计" not in c and "合计" not in c]
                df = df[cols_to_keep]
                # 剔除包含“总计”的行
                df = df[~df.iloc[:, 0].str.contains("总计|合计|Total", na=False)]

                if is_chart_mode:
                    fig, ax = plt.subplots(figsize=(10, 6))
                    # 针对趋势图（只有一行数据，年份在表头）
                    if len(df) == 1:
                        labels = df.columns[1:]
                        vals = [float(re.sub(r'[^\d.]', '', str(v))) for v in df.iloc[0, 1:]]
                        ax.bar(labels, vals, color='#4472C4', width=0.5)
                        for i, v in enumerate(vals):
                            ax.text(i, v, f'{int(v)}', ha='center', va='bottom', fontweight='bold')
                    # 针对分布图（多行数据，如 A-F 级）
                    else:
                        for col in df.columns[1:]:
                            df[col] = pd.to_numeric(df[col].apply(lambda x: re.sub(r'[^\d.]', '', str(x))), errors='coerce').fillna(0)
                        
                        if ("分布" in current_title or "比例" in current_title) and "趋势" not in current_title:
                            # 如果数据是按年份分布的，通常用分组柱状图
                            x = np.arange(len(df))
                            width = 0.8 / (len(df.columns) - 1)
                            for i, col in enumerate(df.columns[1:]):
                                ax.bar(x + i*width, df[col], width, label=col)
                            ax.set_xticks(x + width*(len(df.columns)-2)/2)
                            ax.set_xticklabels(df.iloc[:, 0].astype(str))
                            ax.legend(loc='upper right', fontsize=9)
                        else:
                            # 普通年度变化
                            ax.bar(df.iloc[:, 0].astype(str), df.iloc[:, 1], color='#4472C4', width=0.5)
                    
                    plt.title(current_title, fontsize=12, fontweight='bold', pad=15)
                    plt.tight_layout()
                    buf = io.BytesIO(); plt.savefig(buf, format='png', dpi=180); plt.close(); buf.seek(0)
                    doc.add_picture(buf, width=Inches(5.5))
                    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
                else:
                    table = doc.add_table(rows=len(raw_data), cols=len(raw_data[0]))
                    table.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for i, row_data in enumerate(raw_data):
                        for j, val in enumerate(row_data):
                            cell = table.cell(i, j)
                            cell.text = val
                            run = cell.paragraphs[0].runs[0] if cell.paragraphs[0].runs else cell.paragraphs[0].add_run(val)
                            set_font(run, 10, i==0)
                    set_three_line_table(table)
            except Exception as e:
                print(f"Chart Error: {e}")
                
        table_rows, is_chart_mode, current_title = [], False, ""

    for line in lines:
        l = line.strip()
        if not l: continue
        if l.startswith('#'):
            flush_table()
            level = min(l.count('#'), 3)
            p = doc.add_heading('', level=level)
            run = p.add_run(l.replace('#', '').strip())
            set_font(run, 14, True)
        elif re.match(r'(\*\*?)?[图表]\s?\d+[:：\s]', l):
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
            run = p.add_run(l.replace('**', ''))
            set_font(run, 12)
    flush_table()

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
    run_toc._r.extend([fldChar1, instrText, fldChar2, fldChar3])
    doc.add_page_break()

@app.post("/generate_report_word")
async def generate_report_word(input_data: ReportInput, request: Request):
    try:
        doc = Document()
        # 强制更新域
        element = doc.settings.element
        update_fields = OxmlElement('w:updateFields'); update_fields.set(qn('w:val'), 'true'); element.append(update_fields)
        
        add_page_number(doc); add_cover(doc); add_toc(doc)
        for i in range(1, 9): 
            txt = getattr(input_data, f"ch{i}_text", "")
            if txt and txt.strip(): process_smart(doc, txt)
        
        fname = f"report_{uuid.uuid4().hex[:8]}.docx"
        full_path = os.path.join("static", fname)
        doc.save(full_path)
        return {"file": f"{str(request.base_url).rstrip('/')}/static/{fname}", "status": "success"}
    except Exception as e:
        return {"file": "", "status": "error", "message": str(e)}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8080)
