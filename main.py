import re
import io
import os
import uuid
import pandas as pd
import matplotlib
matplotlib.use('Agg')  # 必须在 pyplot 导入前，适配无界面环境
import matplotlib.pyplot as plt
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
    ch2_text: Optional[str] = ""
    ch3_text: Optional[str] = ""
    ch4_text: Optional[str] = ""
    ch5_text: Optional[str] = ""
    ch6_text: Optional[str] = ""
    ch7_text: Optional[str] = ""

# --- 绘图配置 ---
plt.rcParams['font.sans-serif'] = ['SimHei', 'DejaVu Sans', 'Arial Unicode MS']
plt.rcParams['axes.unicode_minus'] = False

def set_table_border(table):
    """设置学术三线表：顶线、底线 1.5pt，栏目线 0.75pt"""
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

    # 先清除所有边框
    for row in table.rows:
        for cell in row.cells:
            set_cell_border(cell, top={'sz': 0, 'val': 'none'}, bottom={'sz': 0, 'val': 'none'}, 
                            start={'sz': 0, 'val': 'none'}, end={'sz': 0, 'val': 'none'})

    # 设置第一行（顶线）
    for cell in table.rows[0].cells:
        set_cell_border(cell, top={'sz': 12, 'val': 'single', 'color': '000000'})
    
    # 设置第一行底部（栏目线）
    for cell in table.rows[0].cells:
        set_cell_border(cell, bottom={'sz': 6, 'val': 'single', 'color': '000000'})

    # 设置最后一行底部（底线）
    for cell in table.rows[-1].cells:
        set_cell_border(cell, bottom={'sz': 12, 'val': 'single', 'color': '000000'})

def clean_data(text):
    if pd.isna(text): return 0
    val = str(text).replace('$', '').replace(',', '').replace('%', '').replace('*', '').strip()
    try:
        return float(val) if '.' in val else int(val)
    except:
        return val

def md_table_to_df(md_text):
    lines = [l.strip() for l in md_text.strip().split('\n') if '|' in l]
    if len(lines) < 2: return None
    headers = [re.sub(r'[\$\*]', '', c).strip() for c in lines[0].split('|') if c.strip()]
    data = []
    for line in lines:
        if '---' in line or line == lines[0]: continue
        row = [clean_data(c) for c in line.split('|') if c.strip()]
        if len(row) >= len(headers):
            data.append(row[:len(headers)])
    return pd.DataFrame(data, columns=headers) if data else None

def generate_chart(df, title):
    plt.figure(figsize=(9, 5))
    img_stream = io.BytesIO()
    try:
        x_data = df.iloc[:, 0].astype(str)
        y_data = pd.to_numeric(df.iloc[:, 1], errors='coerce').fillna(0)
        
        if any(kw in title for kw in ["分布", "占比", "结构"]):
            plt.pie(y_data, labels=x_data, autopct='%1.1f%%', colors=plt.cm.Pastel1.colors)
        else:
            plt.bar(x_data, y_data, color='royalblue', width=0.5)
            for i, v in enumerate(y_data):
                plt.text(i, v, str(v), ha='center', va='bottom')
        
        plt.title(title, fontsize=12, pad=15)
        plt.tight_layout()
        plt.savefig(img_stream, format='png', dpi=200)
    finally: plt.close()
    img_stream.seek(0)
    return img_stream

def set_run_font(run, size=12, bold=False):
    run.font.size = Pt(size)
    run.font.name = '宋体'
    run.bold = bold
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

def process_content(doc, full_text):
    # 【核心修复】强制在所有标题和表格符号前加换行，解决粘连问题
    full_text = re.sub(r'([^\n])\s*(\|)', r'\1\n\2', full_text)
    full_text = re.sub(r'([^\n])\s*(\*\*?[图表]\s?\d+[:：])', r'\1\n\2', full_text)
    
    # 按照 标题、图表题、表格 进行切割
    parts = re.split(r'(##+ .*?\n|(?:\*\*?)?[图表]\s?\d+[:：].*?\n|(?:\n|^)\|[\s\S]*?\|(?:\n|$))', full_text)
    
    current_fig_title = None
    
    for part in parts:
        if not part or not part.strip(): continue
        p_text = part.strip()
        
        # 1. 标题
        if p_text.startswith('##'):
            h = doc.add_heading('', level=2)
            run = h.add_run(p_text.replace('#','').strip())
            set_run_font(run, 14, True)
            
        # 2. 图表题 (例如 图 1: xxx)
        elif re.match(r'(\*\*?)?[图表]\s?\d+[:：]', p_text):
            current_fig_title = p_text.replace('*', '').strip()
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(current_fig_title)
            set_run_font(run, 11, True)
            
        # 3. 表格/绘图
        elif p_text.startswith('|'):
            df = md_table_to_df(p_text)
            if df is not None:
                if current_fig_title and "图" in current_fig_title:
                    try:
                        img = generate_chart(df, current_fig_title)
                        doc.add_picture(img, width=Inches(5.5))
                        doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    except: pass
                    current_fig_title = None 
                else:
                    # 插入三线表
                    table = doc.add_table(rows=1, cols=len(df.columns))
                    table.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    # 表头
                    for i, col in enumerate(df.columns):
                        cell = table.rows[0].cells[i]
                        cell.text = str(col)
                        set_run_font(cell.paragraphs[0].runs[0], 10.5, True)
                    # 数据
                    for _, row_data in df.iterrows():
                        row_cells = table.add_row().cells
                        for i, val in enumerate(row_data):
                            row_cells[i].text = str(val)
                            set_run_font(row_cells[i].paragraphs[0].runs[0], 10.5)
                    set_table_border(table)
        
        # 4. 普通正文
        else:
            clean_txt = p_text.replace('$', '').replace('***', '')
            if len(clean_txt) > 2:
                p = doc.add_paragraph()
                run = p.add_run(clean_txt)
                set_run_font(run)

@app.post("/generate_report_word")
async def generate_report_word(input_data: ReportInput, request: Request):
    try:
        doc = Document()
        for i in range(2, 8):
            txt = getattr(input_data, f"ch{i}_text", "")
            if txt: process_content(doc, txt)
        
        fname = f"report_{uuid.uuid4().hex[:8]}.docx"
        fpath = os.path.join("static", fname)
        doc.save(fpath)
        return {"status": "success", "file_url": f"{str(request.base_url).rstrip('/')}/static/{fname}"}
    except Exception as e:
        return {"status": "error", "message": str(e)}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=int(os.environ.get("PORT", 8080)))
