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
from typing import Optional, List
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

# --- 绘图与字体 ---
plt.rcParams['font.sans-serif'] = ['SimHei', 'DejaVu Sans', 'Arial Unicode MS']
plt.rcParams['axes.unicode_minus'] = False

def set_table_border(table):
    """标准三线表样式"""
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
    # 表头上下线
    for cell in table.rows[0].cells:
        set_cell_border(cell, top={'sz': 12, 'val': 'single'}, bottom={'sz': 6, 'val': 'single'})
    # 表底线
    for cell in table.rows[-1].cells:
        set_cell_border(cell, bottom={'sz': 12, 'val': 'single'})

def clean_val(x):
    s = str(x).replace(',', '').replace('%', '').replace('*', '').strip()
    try:
        return float(s) if '.' in s else int(s)
    except:
        return s

def generate_chart(df, title):
    """将提取的表格转为柱状图或饼图"""
    plt.figure(figsize=(8, 4.5))
    try:
        # 强制将第二列及以后转为数字
        x_data = df.iloc[:, 0].astype(str)
        y_data = pd.to_numeric(df.iloc[:, 1].apply(clean_val), errors='coerce').fillna(0)
        
        if any(k in title for k in ["占比", "分布", "结构"]):
            plt.pie(y_data, labels=x_data, autopct='%1.1f%%', colors=plt.cm.Pastel1.colors)
        else:
            bars = plt.bar(x_data, y_data, color='#4472C4', width=0.5)
            for bar in bars:
                plt.text(bar.get_x() + bar.get_width()/2, bar.get_height(), f'{bar.get_height()}', ha='center', va='bottom')
        
        plt.title(title, fontsize=11)
        plt.tight_layout()
        img_stream = io.BytesIO()
        plt.savefig(img_stream, format='png', dpi=200)
        plt.close()
        img_stream.seek(0)
        return img_stream
    except Exception as e:
        print(f"绘图报错: {e}")
        return None

def set_font(run, size=12, bold=False):
    run.font.size = Pt(size)
    run.font.name = '宋体'
    run.bold = bold
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

def parse_and_write(doc, text):
    """逐行扫描状态机：修复标题识别、表格识别和换行问题"""
    # 统一换行符并分割行
    lines = text.replace('\r\n', '\n').replace('\r', '\n').split('\n')
    table_buffer = []
    current_title = ""
    pending_text = ""  # 用于累积非表格的文本段落

    def flush_pending_text():
        """将累积的文本写入段落"""
        nonlocal pending_text
        if pending_text.strip():
            p = doc.add_paragraph()
            run = p.add_run(pending_text.strip())
            set_font(run, 12)
            pending_text = ""

    def flush_table():
        nonlocal table_buffer, current_title
        if not table_buffer: return
        
        # 解析表格
        rows = []
        for lb in table_buffer:
            cells = [c.strip() for c in lb.split('|') if c.strip()]
            if cells and not all('---' in c for c in cells):
                rows.append(cells)
        
        if len(rows) > 1:
            df = pd.DataFrame(rows[1:], columns=rows[0])
            # 判断是画图还是画表
            if current_title and "图" in current_title:
                img = generate_chart(df, current_title)
                if img:
                    doc.add_picture(img, width=Inches(5.5))
                    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                # 即使没有标题也画个三线表
                table = doc.add_table(rows=len(rows), cols=len(rows[0]))
                table.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for r_idx, row_data in enumerate(rows):
                    for c_idx, val in enumerate(row_data):
                        cell = table.cell(r_idx, c_idx)
                        cell.text = str(val)
                        set_font(cell.paragraphs[0].runs[0], 10.5, r_idx==0)
                set_table_border(table)
        
        table_buffer = []
        current_title = ""

    for line in lines:
        clean_line = line.strip()
        
        # 空行：刷新文本和表格，实现段落分隔
        if not clean_line:
            flush_pending_text()
            flush_table()
            continue
        
        # 1. 识别Markdown标题（# 开头）
        if clean_line.startswith('#'):
            flush_pending_text()
            flush_table()
            # 计算标题级别
            level = 1
            while level <= 3 and clean_line.startswith('#'*level):
                level += 1
            level -= 1
            # 添加标题
            title_text = clean_line.lstrip('#').strip()
            h = doc.add_heading('', level=level)
            run = h.add_run(title_text)
            set_font(run, 16 - level*2, True)
            continue
        
        # 2. 识别图表标题 (例如: 图 1：... 或 **表 1：...**)
        if re.match(r'(\*\*?)?[图表]\s?\d+[:：]', clean_line):
            flush_pending_text()
            flush_table()
            current_title = clean_line.replace('*', '').strip()
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(current_title)
            set_font(run, 11, True)
            continue
        
        # 3. 识别表格行（以 | 开头）
        if clean_line.startswith('|'):
            flush_pending_text()
            table_buffer.append(clean_line)
            continue
        
        # 4. 普通正文：累积文本，避免逐行添加导致的换行问题
        pending_text += clean_line + " "

    # 循环结束后，刷新所有剩余内容
    flush_pending_text()
    flush_table()

@app.post("/generate_report_word")
async def generate_report_word(input_data: ReportInput, request: Request):
    try:
        doc = Document()
        # 设置默认字体
        doc.styles['Normal'].font.name = '宋体'
        doc.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
        doc.styles['Normal'].font.size = Pt(12)
        
        # 依次处理每一章
        chapters = [input_data.ch2_text, input_data.ch3_text, input_data.ch4_text, 
                    input_data.ch5_text, input_data.ch6_text, input_data.ch7_text]
        for ch in chapters:
            if ch and len(ch.strip()) > 5:
                parse_and_write(doc, ch)
        
        file_id = uuid.uuid4().hex[:8]
        file_name = f"report_{file_id}.docx"
        file_path = os.path.join("static", file_name)
        doc.save(file_path)
        
        return {
            "status": "success",
            "file_url": f"{str(request.base_url).rstrip('/')}/static/{file_name}"
        }
    except Exception as e:
        import traceback
        traceback.print_exc()
        return {"status": "error", "message": str(e)}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=int(os.environ.get("PORT", 8080)))
