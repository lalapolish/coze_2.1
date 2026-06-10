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

# 绘图配置
plt.rcParams['font.sans-serif'] = ['DejaVu Sans', 'Arial Unicode MS']
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
    """标准学术三线表"""
    for row_idx, row in enumerate(table.rows):
        for cell in row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            for border_name in ['top', 'bottom', 'left', 'right']:
                edge = OxmlElement(f'w:{border_name}')
                # 顶线和底线 1.5磅(sz=12)，中间线 0.75磅(sz=6)
                if border_name == 'top' and row_idx == 0:
                    edge.set(qn('w:val'), 'single')
                    edge.set(qn('w:sz'), '12')
                elif border_name == 'bottom' and row_idx == 0:
                    edge.set(qn('w:val'), 'single')
                    edge.set(qn('w:sz'), '6')
                elif border_name == 'bottom' and row_idx == len(table.rows) - 1:
                    edge.set(qn('w:val'), 'single')
                    edge.set(qn('w:sz'), '12')
                else:
                    edge.set(qn('w:val'), 'none')
                tcPr.append(edge)

def process_smart(doc, text):
    # 1. 预处理：删掉干扰符号
    text = text.replace('$$', '').replace('\r', '')
    lines = text.split('\n')
    
    table_rows = []
    current_title = ""
    is_chart_mode = False

    def flush_table():
        nonlocal table_rows, current_title, is_chart_mode
        if not table_rows: return
        
        # 提取有效数据
        data = []
        for r in table_rows:
            cells = [c.strip() for c in r.split('|') if c.strip()]
            if cells and '---' not in r: data.append(cells)
        
        if len(data) >= 2:
            try:
                if is_chart_mode:
                    # 绘图逻辑
                    df = pd.DataFrame(data[1:], columns=data[0])
                    plt.figure(figsize=(8, 4))
                    x = df.iloc[:, 0].astype(str)
                    # 转换数字，处理逗号
                    y = pd.to_numeric(df.iloc[:, 1].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
                    
                    if "比例" in current_title or "分布" in current_title and len(x) < 10:
                        plt.pie(y, labels=x, autopct='%1.1f%%', colors=plt.cm.Pastel1.colors)
                    else:
                        bars = plt.bar(x, y, color='#4472C4', width=0.5)
                        for bar in bars:
                            plt.text(bar.get_x() + bar.get_width()/2, bar.get_height(), f'{int(bar.get_height())}', ha='center', va='bottom', fontsize=8)
                    
                    plt.title(current_title, fontsize=10)
                    plt.tight_layout()
                    
                    buf = io.BytesIO()
                    plt.savefig(buf, format='png', dpi=150)
                    plt.close()
                    buf.seek(0)
                    doc.add_picture(buf, width=Inches(5))
                    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
                else:
                    # 三线表逻辑
                    table = doc.add_table(rows=len(data), cols=len(data[0]))
                    table.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for i, row_data in enumerate(data):
                        for j, val in enumerate(row_data):
                            cell = table.cell(i, j)
                            cell.text = val
                            set_font(cell.paragraphs[0].runs[0], 10, i==0)
                    set_table_border(table)
            except Exception as e:
                print(f"表格转换错误: {e}")
        
        table_rows = []
        is_chart_mode = False
        current_title = ""

    for line in lines:
        l = line.strip()
        if not l: continue

        # A. 识别标题 (##)
        if l.startswith('#'):
            flush_table()
            level = min(l.count('#'), 3)
            p = doc.add_heading('', level=level)
            run = p.add_run(l.replace('#', '').strip())
            set_font(run, 16 - level*2, True)

        # B. 识别图表说明 (图1: xxx)
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

        # D. 普通文本
        else:
            if table_rows: flush_table()
            p = doc.add_paragraph()
            run = p.add_run(l.replace('**', ''))
            set_font(run, 12)

    flush_table()

@app.post("/generate_report_word")
async def generate_report_word(input_data: ReportInput, request: Request):
    try:
        doc = Document()
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
