import re
import io
import os
import uuid
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
    ch2_text: Optional[str] = ""
    ch3_text: Optional[str] = ""
    ch4_text: Optional[str] = ""
    ch5_text: Optional[str] = ""
    ch6_text: Optional[str] = ""
    ch7_text: Optional[str] = ""

# 绘图配置 (尝试适配 Linux)
plt.rcParams['font.sans-serif'] = ['DejaVu Sans', 'Liberation Sans', 'Arial']
plt.rcParams['axes.unicode_minus'] = False

def set_font(run, size=12, bold=False):
    run.font.size = Pt(size)
    run.font.name = '宋体'
    run.bold = bold
    rPr = run._element.get_or_add_rPr()
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:eastAsia'), '宋体')
    rPr.append(rFonts)

def set_full_border(table):
    """为表格添加全边框"""
    for row in table.rows:
        for cell in row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            for border_name in ['top', 'bottom', 'left', 'right']:
                edge = OxmlElement(f'w:{border_name}')
                edge.set(qn('w:val'), 'single')
                edge.set(qn('w:sz'), '4') # 0.5 磅
                edge.set(qn('w:space'), '0')
                edge.set(qn('w:color'), 'auto')
                tcPr.append(edge)

def process_smart(doc, text):
    # 预处理：删除干扰符
    text = text.replace('$$', '').replace('\r', '')
    lines = text.split('\n')
    
    table_rows = []
    current_title = ""
    is_chart_mode = False

    def flush_table():
        nonlocal table_rows, current_title, is_chart_mode
        if not table_rows: return
        
        data = []
        for r in table_rows:
            cells = [c.strip() for c in r.split('|') if c.strip()]
            if cells and '---' not in r: data.append(cells)
        
        if len(data) >= 2:
            try:
                if is_chart_mode:
                    # 绘图逻辑
                    df = pd.DataFrame(data[1:], columns=data[0])
                    plt.figure(figsize=(8, 4.5))
                    x = df.iloc[:, 0].astype(str)
                    y = pd.to_numeric(df.iloc[:, 1].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
                    
                    if "比例" in current_title or "分布" in current_title:
                        # 饼图：标注名称和百分比
                        plt.pie(y, labels=x, autopct='%1.1f%%', startangle=140, colors=plt.cm.Paired.colors)
                    else:
                        # 柱状图
                        bars = plt.bar(x, y, color='#4472C4', width=0.5)
                        for bar in bars:
                            plt.text(bar.get_x() + bar.get_width()/2, bar.get_height(), 
                                     f'{int(bar.get_height())}', ha='center', va='bottom', fontsize=9)
                    
                    # 关键修改：图片内部不写标题，避免黑框，标题由 Word 正文提供
                    plt.tight_layout()
                    
                    buf = io.BytesIO()
                    plt.savefig(buf, format='png', dpi=150)
                    plt.close()
                    buf.seek(0)
                    doc.add_picture(buf, width=Inches(5.2))
                    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
                else:
                    # 全框表格逻辑
                    table = doc.add_table(rows=len(data), cols=len(data[0]))
                    table.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for i, row_data in enumerate(data):
                        for j, val in enumerate(row_data):
                            cell = table.cell(i, j)
                            cell.text = val
                            set_font(cell.paragraphs[0].runs[0], 10.5, i==0)
                    set_full_border(table)
            except Exception as e:
                print(f"表格转换错误: {e}")
        
        table_rows = []
        is_chart_mode = False
        current_title = ""

    for line in lines:
        l = line.strip()
        if not l: continue

        # A. 识别标题 (##) - 不缩进
        if l.startswith('#'):
            flush_table()
            level = min(l.count('#'), 3)
            p = doc.add_heading('', level=level)
            run = p.add_run(l.replace('#', '').strip())
            set_font(run, 15 - level, True)

        # B. 识别图表说明 - 居中且不缩进
        elif re.match(r'(\*\*?)?[图表]\s?\d+[:：]', l):
            flush_table()
            current_title = l.replace('*', '').strip()
            is_chart_mode = "图" in current_title
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(current_title)
            set_font(run, 11, True)

        # C. 识别表格行
        elif l.startswith('|'):
            table_rows.append(l)

        # D. 普通文本 - 设置首行缩进 2 字符
        else:
            if table_rows: flush_table()
            p = doc.add_paragraph()
            # 设置首行缩进: 12pt字体下，2字符 = 24磅
            p.paragraph_format.first_line_indent = Pt(24) 
            run = p.add_run(l.replace('**', ''))
            set_font(run, 12)

    flush_table()

@app.post("/generate_report_word")
async def generate_report_word(input_data: ReportInput, request: Request):
    try:
        doc = Document()
        # 设置默认节属性（如需设置页边距可在此处）
        for i in range(2, 8):
            content = getattr(input_data, f"ch{i}_text", "")
            if content: process_smart(doc, content)
        
        fname = f"report_{uuid.uuid4().hex[:8]}.docx"
        path = os.path.join("static", fname)
        doc.save(path)
        return {"status": "success", "file_url": f"{str(request.base_url).rstrip('/')}/static/{fname}"}
    except Exception as e:
        return {"status": "error", "message": str(e)}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8080)
