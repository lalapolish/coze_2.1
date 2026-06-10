import re
import io
import os
import uuid
import pandas as pd
import matplotlib.pyplot as plt
from fastapi import FastAPI, Request
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
from typing import Optional
from docx import Document
from docx.shared import Pt, Inches
from docx.oxml.ns import qn
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

# --- 绘图配置 (增加兼容性) ---
try:
    plt.rcParams['font.sans-serif'] = ['SimHei', 'DejaVu Sans', 'Arial Unicode MS'] 
except:
    pass
plt.rcParams['axes.unicode_minus'] = False
BLUE_COLOR = 'royalblue'
GROUPED_COLORS = ['#4472C4', '#ED7D31', '#FFC000', '#70AD47', '#25B6C7']

def clean_data(text):
    if pd.isna(text): return 0
    val = str(text).replace('$', '').replace(',', '').replace('%', '').replace('*', '').strip()
    try:
        if '.' in val: return float(val)
        return int(val)
    except:
        return val

def md_table_to_df(md_text):
    # 过滤掉空行，只保留带 | 的行
    lines = [l.strip() for l in md_text.strip().split('\n') if '|' in l]
    if len(lines) < 2: return None
    # 提取表头
    headers = [re.sub(r'[\$\*]', '', c).strip() for c in lines[0].split('|') if c.strip()]
    data = []
    # 寻找数据行（跳过表头和分隔行 ---|---）
    for line in lines:
        if '---' in line or line == lines[0]: continue
        row = [clean_data(c) for c in line.split('|') if c.strip()]
        if len(row) >= len(headers):
            data.append(row[:len(headers)])
    return pd.DataFrame(data, columns=headers) if data else None

def generate_chart(df, title, fig_no):
    plt.figure(figsize=(9, 5))
    img_stream = io.BytesIO()
    try:
        # 简单逻辑：如果第一列是年份或类别，第二列是数值
        x_data = df.iloc[:, 0].astype(str)
        y_data = pd.to_numeric(df.iloc[:, 1], errors='coerce').fillna(0)
        
        if fig_no in [4, 5] or "分布" in title or "占比" in title:
            plt.pie(y_data, labels=x_data, autopct='%1.1f%%', colors=plt.cm.Pastel1.colors)
        else:
            bars = plt.bar(x_data, y_data, color=BLUE_COLOR, width=0.6)
            for bar in bars:
                plt.text(bar.get_x() + bar.get_width()/2, bar.get_height(), f'{bar.get_height()}', ha='center', va='bottom')
        
        plt.title(title, fontsize=12, pad=15)
        plt.xticks(rotation=15 if len(x_data) > 5 else 0)
        plt.tight_layout()
        plt.savefig(img_stream, format='png', dpi=200)
    finally: 
        plt.close()
    img_stream.seek(0)
    return img_stream

def set_style(obj, is_title=False):
    if hasattr(obj, 'runs'):
        for run in obj.runs:
            run.font.size = Pt(14 if is_title else 12)
            run.font.name = '宋体'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

def process_content(doc, full_text):
    # --- 核心优化：更强大的切分正则 ---
    # 匹配 ##标题、图表标记（支持星号可选、冒号可选）、Markdown表格
    parts = re.split(r'(##+ .*?|(?:\*\*?)?[图表]\s?\d+[:：].*?(?:\*\*?)?|(?:\n|^)\|.*\|(?:\n|$))', full_text, flags=re.S)
    
    current_fig_title = None
    
    for part in parts:
        if not part or not part.strip(): continue
        part_s = part.strip()
        
        # 1. 处理标题
        if part_s.startswith('##'):
            h = doc.add_heading(part_s.replace('#','').strip(), level=2)
            set_style(h, True)
            
        # 2. 处理图表标记 (例如: 图 1: xxx)
        elif re.match(r'(\*\*?)?[图表]\s?\d+[:：]', part_s):
            current_fig_title = part_s.replace('*', '')
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(current_fig_title)
            run.bold = True
            set_style(p)
            
        # 3. 处理表格内容
        elif part_s.startswith('|'):
            df = md_table_to_df(part_s)
            if df is not None:
                if current_fig_title and ("图" in current_fig_title):
                    try:
                        fig_no_match = re.search(r'\d+', current_fig_title)
                        fig_no = int(fig_no_match.group()) if fig_no_match else 0
                        img = generate_chart(df, current_fig_title, fig_no)
                        doc.add_picture(img, width=Inches(5.5))
                        doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    except Exception as e:
                        print(f"绘图失败: {e}")
                    current_fig_title = None 
                else:
                    # 插入普通 Word 表格
                    table = doc.add_table(rows=1, cols=len(df.columns))
                    table.style = 'Table Grid'
                    for i, col in enumerate(df.columns): table.rows[0].cells[i].text = str(col)
                    for _, row in df.iterrows():
                        row_cells = table.add_row().cells
                        for i, val in enumerate(row): row_cells[i].text = str(val)
            
        # 4. 处理普通文本
        else:
            p = doc.add_paragraph(part_s.replace('$', ''))
            set_style(p)

@app.post("/generate_report_word")
async def generate_report_word(input_data: ReportInput, request: Request):
    try:
        doc = Document()
        # 依次处理 2-7 章
        for i in range(2, 8):
            content = getattr(input_data, f"ch{i}_text", "")
            if content and len(content.strip()) > 5:
                process_content(doc, content)
        
        file_id = uuid.uuid4().hex[:8]
        file_name = f"report_{file_id}.docx"
        file_path = os.path.join("static", file_name)
        doc.save(file_path)
        
        base_url = str(request.base_url).rstrip('/')
        return {
            "status": "success",
            "file_url": f"{base_url}/static/{file_name}",
            "message": "文档生成成功"
        }
    except Exception as e:
        return {"status": "error", "message": f"生成失败: {str(e)}"}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=int(os.environ.get("PORT", 8080)))
