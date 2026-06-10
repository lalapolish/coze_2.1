import re
import io
import os
import uuid
import pandas as pd
import matplotlib
matplotlib.use('Agg') # 必须在最前面
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

# --- 绘图配置 (适配 Linux 环境防止乱码) ---
plt.rcParams['font.sans-serif'] = ['DejaVu Sans', 'Arial Unicode MS'] # 使用更通用的字体
plt.rcParams['axes.unicode_minus'] = False

def clean_val(x):
    """极其严格的数字清洗，处理 $$ 符号"""
    if pd.isna(x): return 0
    # 去掉所有非数字和非小数点的符号
    s = str(x).replace('$', '').replace(',', '').replace('%', '').strip()
    try:
        return float(s) if '.' in s else int(s)
    except:
        return 0

def md_table_to_df(md_text):
    """将单个 Markdown 表格转为 DataFrame"""
    lines = [l.strip() for l in md_text.strip().split('\n') if '|' in l]
    if len(lines) < 2: return None
    
    # 提取表头并去掉 $$
    headers = [c.strip().replace('$', '') for c in lines[0].split('|') if c.strip()]
    
    data = []
    for line in lines:
        if '---' in line or line == lines[0]: continue
        # 提取单元格并去掉 $$
        row = [c.strip().replace('$', '') for c in line.split('|') if c.strip()]
        if len(row) >= len(headers):
            data.append(row[:len(headers)])
            
    if not data: return None
    return pd.DataFrame(data, columns=headers)

def generate_chart(df, title):
    """画图：强制只取前两列，防止多表合并错误"""
    plt.figure(figsize=(9, 4.5))
    try:
        # 强制只取前两列，第一列为类别，第二列为数值
        x_label = df.columns[0]
        y_label = df.columns[1]
        
        x_data = df.iloc[:, 0].astype(str)
        y_data = df.iloc[:, 1].apply(clean_val)
        
        if "分布" in title or "占比" in title:
            plt.pie(y_data, labels=x_data, autopct='%1.1f%%', colors=plt.cm.Pastel1.colors)
        else:
            bars = plt.bar(x_data, y_data, color='#4472C4', width=0.5)
            for bar in bars:
                plt.text(bar.get_x() + bar.get_width()/2, bar.get_height(), 
                         f'{int(bar.get_height())}', ha='center', va='bottom', fontsize=9)
        
        # 简单处理标题中的中文字体显示（Linux环境下可能仍有挑战，建议用英文或去除特殊字符）
        plt.title(title, fontsize=10)
        plt.tight_layout()
        
        img_stream = io.BytesIO()
        plt.savefig(img_stream, format='png', dpi=180)
        plt.close()
        img_stream.seek(0)
        return img_stream
    except Exception as e:
        print(f"绘图出错: {e}")
        return None

def set_font(run, size=12, bold=False):
    run.font.size = Pt(size)
    run.font.name = '宋体'
    run.bold = bold
    # 强制东亚字体渲染
    r = run._element
    rPr = r.get_or_add_rPr()
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:eastAsia'), '宋体')
    rPr.append(rFonts)

def process_content(doc, full_text):
    # 1. 先把干扰的 $$ 去掉，防止正文数字粘连
    full_text = full_text.replace('$$', '')
    
    # 2. 使用非贪婪切分：按标题、图表标记、表格块切分
    # 核心：表格块匹配改为非贪婪，避免把两个表格连在一起
    parts = re.split(r'(###+ .*?\n|(?:\*\*?)?[图表]\s?\d+[:：].*?\n|\|(?:.*?\|)+\n)', full_text)
    
    current_fig_title = None
    
    for part in parts:
        if not part or not part.strip(): continue
        p_text = part.strip()
        
        # 处理标题
        if p_text.startswith('##'):
            level = p_text.count('#')
            h = doc.add_heading('', level=min(level, 3))
            run = h.add_run(p_text.replace('#','').strip())
            set_font(run, 14 if level==2 else 12, True)
            
        # 处理图表题
        elif re.match(r'(\*\*?)?[图表]\s?\d+[:：]', p_text):
            current_fig_title = p_text.replace('*', '').strip()
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(current_fig_title)
            set_font(run, 11, True)
            
        # 处理表格内容
        elif p_text.startswith('|'):
            df = md_table_to_df(p_text)
            if df is not None:
                if current_fig_title and "图" in current_fig_title:
                    img = generate_chart(df, current_fig_title)
                    if img:
                        doc.add_picture(img, width=Inches(5.5))
                        doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    current_fig_title = None 
                else:
                    # 插入普通 Word 表格
                    table = doc.add_table(rows=len(df)+1, cols=len(df.columns))
                    table.style = 'Table Grid'
                    # 表头
                    for i, col in enumerate(df.columns):
                        cell = table.cell(0, i)
                        cell.text = str(col)
                        set_font(cell.paragraphs[0].runs[0], 10, True)
                    # 数据
                    for r_idx, row in enumerate(df.values):
                        for c_idx, val in enumerate(row):
                            cell = table.cell(r_idx+1, c_idx)
                            cell.text = str(val)
                            set_font(cell.paragraphs[0].runs[0], 10)
        
        # 普通正文
        else:
            p = doc.add_paragraph()
            run = p.add_run(p_text.replace('**', ''))
            set_font(run, 12)

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
