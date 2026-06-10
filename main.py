import re
import io
import os
import uuid
import pandas as pd
import matplotlib.pyplot as plt
from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import StreamingResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
from typing import Optional, List
from docx import Document
from docx.shared import Pt, Inches
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH

# --- 初始化 FastAPI ---
app = FastAPI(openapi_version="3.0.0")

# 创建一个 static 文件夹来存放生成的 Word 文件
if not os.path.exists("static"):
    os.makedirs("static")

# 将 static 文件夹映射为可以通过 URL 访问的路径
app.mount("/static", StaticFiles(directory="static"), name="static")

# --- 输入模型定义 ---
class ReportInput(BaseModel):
    ch2_text: Optional[str] = ""
    ch3_text: Optional[str] = ""
    ch4_text: Optional[str] = ""
    ch5_text: Optional[str] = ""
    ch6_text: Optional[str] = ""
    ch7_text: Optional[str] = ""

# --- 绘图全局配置 ---
plt.rcParams['font.sans-serif'] = ['SimHei'] 
plt.rcParams['axes.unicode_minus'] = False
BLUE_COLOR = 'royalblue'
GROUPED_COLORS = ['#4472C4', '#ED7D31', '#FFC000', '#70AD47', '#25B6C7']

def clean_data(text):
    if pd.isna(text): return 0
    clean_val = str(text).replace('$', '').replace(',', '').replace('%', '').strip()
    try:
        return float(clean_val) if '.' in clean_val else int(clean_val)
    except:
        return clean_val

def md_table_to_df(md_text):
    lines = [line.strip() for line in md_text.strip().split('\n') if '|' in line]
    if len(lines) < 2: return None
    headers = [re.sub(r'[\$\*]', '', c).strip() for c in lines[0].split('|') if c.strip()]
    data = []
    for line in lines[2:]:
        row = [clean_data(c) for c in line.split('|') if c.strip()]
        if len(row) == len(headers):
            data.append(row)
    return pd.DataFrame(data, columns=headers)

def generate_chart(df, title, fig_no):
    plt.figure(figsize=(10, 6))
    img_stream = io.BytesIO()
    try:
        if fig_no in [2, 12]:
            df.set_index(df.columns[0]).plot(kind='bar', color=GROUPED_COLORS, ax=plt.gca(), width=0.8)
            plt.legend(bbox_to_anchor=(0.5, -0.15), loc='upper center', ncol=3, frameon=False)
        elif fig_no in [4, 5]:
            plt.pie(df.iloc[:, 1], labels=df.iloc[:, 0], autopct='%1.1f%%', colors=plt.cm.Pastel1.colors)
        else:
            bars = plt.bar(df.iloc[:, 0].astype(str), df.iloc[:, 1], color=BLUE_COLOR)
            for bar in bars:
                plt.text(bar.get_x() + bar.get_width()/2, bar.get_height(), f'{bar.get_height()}', ha='center', va='bottom')
        plt.title(title, fontsize=14, pad=20)
        plt.tight_layout()
        plt.savefig(img_stream, format='png', dpi=300)
    finally: 
        plt.close()
    img_stream.seek(0)
    return img_stream

def set_style(obj):
    if hasattr(obj, 'runs'):
        for run in obj.runs:
            run.font.size = Pt(12)
            run.font.name = '宋体'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

def process_content(doc, full_text):
    parts = re.split(r'(\*\*图 \d+：.*?\*\*|\*\*表 \d+：.*?\*\*|## .*?|### .*?)', full_text)
    current_fig_title = None
    for part in parts:
        part = part.strip()
        if not part: continue
        if part.startswith('##'):
            h = doc.add_heading(part.replace('#','').strip(), level=2)
            set_style(h)
        elif "**图" in part:
            current_fig_title = part.strip("*")
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(current_fig_title)
            run.bold = True
            set_style(p)
        elif part.startswith('|'):
            df = md_table_to_df(part)
            if df is not None:
                if current_fig_title:
                    fig_no_match = re.search(r'图 (\d+)', current_fig_title)
                    fig_no = int(fig_no_match.group(1)) if fig_no_match else 0
                    img = generate_chart(df, current_fig_title, fig_no)
                    doc.add_picture(img, width=Inches(5.8))
                    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    current_fig_title = None
                else:
                    table = doc.add_table(rows=1, cols=len(df.columns))
                    table.style = 'Table Grid'
                    for i, col in enumerate(df.columns): table.rows[0].cells[i].text = str(col)
                    for _, row in df.iterrows():
                        row_cells = table.add_row().cells
                        for i, val in enumerate(row): row_cells[i].text = str(val)
        else:
            p = doc.add_paragraph(part.replace('$', ''))
            set_style(p)

@app.post("/generate_report_word")
async def generate_report_word(input_data: ReportInput, request: Request):
    try:
        doc = Document()
        chapters = [input_data.ch2_text, input_data.ch3_text, input_data.ch4_text, input_data.ch5_text, input_data.ch6_text, input_data.ch7_text]
        for content in chapters:
            if content and len(content.strip()) > 0:
                process_content(doc, content)
        
        # --- 关键修改：保存文件到本地并返回链接 ---
        file_id = uuid.uuid4().hex[:8]
        file_name = f"report_{file_id}.docx"
        file_path = os.path.join("static", file_name)
        doc.save(file_path)
        
        # 构建下载链接 (根据你的 Render 域名)
        # 你的域名是 https://coze-2-1.onrender.com
        base_url = str(request.base_url).rstrip('/')
        download_url = f"{base_url}/static/{file_name}"
        
        return {
            "status": "success",
            "file_url": download_url,
            "message": "Word文档已生成，请通过链接下载"
        }
    except Exception as e:
        return {"status": "error", "message": str(e)}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=int(os.environ.get("PORT", 8080)))
