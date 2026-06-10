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

plt.rcParams['font.sans-serif'] = ['SimHei', 'DejaVu Sans', 'Arial Unicode MS']
plt.rcParams['axes.unicode_minus'] = False

def set_table_border(table):
    """真正符合中文学术规范的三线表"""
    def set_cell_border(cell, **kwargs):
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        for edge in ('top', 'start', 'bottom', 'end'):
            if edge in kwargs:
                tag = 'w:{}'.format(edge)
                element = tcPr.find(qn(tag))
                if element is None:
                    element = OxmlElement(tag)
                    tcPr.append(element)
                for key, val in kwargs[edge].items():
                    element.set(qn('w:{}'.format(key)), str(val))

    for row in table.rows:
        for cell in row.cells:
            # 清除所有边框
            set_cell_border(cell, top={'sz': 0, 'val': 'none'}, bottom={'sz': 0, 'val': 'none'}, 
                            start={'sz': 0, 'val': 'none'}, end={'sz': 0, 'val': 'none'})
    
    # 设置顶线（1.5磅）
    for cell in table.rows[0].cells:
        set_cell_border(cell, top={'sz': 12, 'val': 'single', 'color': '000000'})
    # 设置栏目线（0.75磅）
    for cell in table.rows[0].cells:
        set_cell_border(cell, bottom={'sz': 6, 'val': 'single', 'color': '000000'})
    # 设置底线（1.5磅）
    for cell in table.rows[-1].cells:
        set_cell_border(cell, bottom={'sz': 12, 'val': 'single', 'color': '000000'})

def set_run_style(run, size=12, bold=False):
    run.font.size = Pt(size)
    run.font.name = '宋体'
    run.bold = bold
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

def smart_process(doc, text):
    """
    终极解析逻辑：
    1. 强行修复大模型输出的“粘连”文本
    2. 自动区分标题、正文、图表和表格
    """
    # --- 第一步：强行纠正粘连 ---
    # 在所有标题 (##)、图表说明 (**图)、表格 (|) 前强行加换行符
    text = re.sub(r'(#+ )', r'\n\1', text)
    text = re.sub(r'(\*\*?[图表]\s?\d+[:：])', r'\n\1', text)
    text = re.sub(r'([^|\n])\s*(\|)', r'\1\n\2', text)
    
    lines = text.split('\n')
    table_buffer = []
    last_title = ""

    def flush_table():
        nonlocal table_buffer, last_title
        if not table_buffer: return
        
        # 提取有效数据行
        data_rows = []
        for line in table_buffer:
            cells = [c.strip() for c in line.split('|') if c.strip()]
            if cells and not all('-' in c for c in cells): # 排除分隔线
                data_rows.append(cells)
        
        if len(data_rows) >= 2:
            try:
                # 判断是画图还是画表
                if last_title and "图" in last_title:
                    df = pd.DataFrame(data_rows[1:], columns=data_rows[0])
                    # 转换数值
                    x = df.iloc[:, 0].astype(str)
                    y = pd.to_numeric(df.iloc[:, 1].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
                    
                    plt.figure(figsize=(8, 4))
                    plt.bar(x, y, color='#4472C4', width=0.5)
                    plt.title(last_title, fontsize=10)
                    plt.tight_layout()
                    
                    img_stream = io.BytesIO()
                    plt.savefig(img_stream, format='png', dpi=150)
                    plt.close()
                    img_stream.seek(0)
                    doc.add_picture(img_stream, width=Inches(5))
                    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
                else:
                    # 插入学术三线表
                    table = doc.add_table(rows=len(data_rows), cols=len(data_rows[0]))
                    table.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for i, row_cells in enumerate(data_rows):
                        for j, val in enumerate(row_cells):
                            cell = table.cell(i, j)
                            cell.text = str(val)
                            set_run_style(cell.paragraphs[0].runs[0], 10.5, i==0)
                    set_table_border(table)
            except Exception as e:
                print(f"处理图表出错: {e}")
        
        table_buffer = []
        last_title = ""

    for line in lines:
        row = line.strip()
        if not row: continue

        # 1. 处理标题
        if row.startswith('#'):
            flush_table()
            level = min(row.count('#'), 3)
            p = doc.add_heading('', level=level)
            run = p.add_run(row.replace('#', '').strip())
            set_run_style(run, 16 - level*2, True)

        # 2. 处理图表说明
        elif re.match(r'(\*\*?)?[图表]\s?\d+[:：]', row):
            flush_table()
            last_title = row.replace('*', '').strip()
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(last_title)
            set_run_style(run, 11, True)

        # 3. 处理表格行
        elif row.startswith('|'):
            table_buffer.append(row)

        # 4. 处理普通正文
        else:
            if table_buffer: flush_table()
            p = doc.add_paragraph()
            run = p.add_run(row.replace('$', '').replace('**', ''))
            set_run_style(run, 12)

    flush_table()

@app.post("/generate_report_word")
async def generate_report_word(input_data: ReportInput, request: Request):
    try:
        doc = Document()
        # 处理所有章节内容
        for i in range(2, 8):
            content = getattr(input_data, f"ch{i}_text", "")
            if content: smart_process(doc, content)
        
        fname = f"report_{uuid.uuid4().hex[:8]}.docx"
        path = os.path.join("static", fname)
        doc.save(path)
        return {"status": "success", "file_url": f"{str(request.base_url).rstrip('/')}/static/{fname}"}
    except Exception as e:
        return {"status": "error", "message": str(e)}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=int(os.environ.get("PORT", 8080)))
