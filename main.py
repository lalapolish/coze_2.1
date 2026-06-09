import re
import io
import os
import pandas as pd
import matplotlib.pyplot as plt
from fastapi import FastAPI, HTTPException
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
from typing import Optional
from docx import Document
from docx.shared import Pt, Inches
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
import traceback

# --- 初始化 FastAPI ---
app = FastAPI(openapi_version="3.0.2")

# --- 输入模型定义 ---
class ReportInput(BaseModel):
    ch2_text: Optional[str] = ""
    ch3_text: Optional[str] = ""
    ch4_text: Optional[str] = ""
    ch5_text: Optional[str] = ""
    ch6_text: Optional[str] = ""
    ch7_text: Optional[str] = ""

# =================================================================
# 核心逻辑：绘图与表格处理
# =================================================================

# 绘图全局样式设置
plt.rcParams['font.sans-serif'] = ['SimHei']  # 注意：在 Linux 环境部署需确保安装了黑体
plt.rcParams['axes.unicode_minus'] = False
BLUE_COLOR = 'royalblue'
GROUPED_COLORS = ['#4472C4', '#ED7D31', '#FFC000', '#70AD47', '#25B6C7']

def clean_data(text):
    if pd.isna(text): return 0
    clean_val = str(text).replace('$', '').replace(',', '').replace('%', '').strip()
    try:
        if '.' in clean_val:
            return float(clean_val)
        return int(clean_val)
    except:
        return clean_val

def md_table_to_df(md_text):
    lines = [line.strip() for line in md_text.strip().split('\n') if '|' in line]
    if len(lines) < 2: return None
    # 提取表头并清理 $ 符号
    headers = [re.sub(r'[\$\*]', '', c).strip() for c in lines[0].split('|') if c.strip()]
    data = []
    # 跳过表头和分割线
    for line in lines[2:]:
        row = [clean_data(c) for c in line.split('|') if c.strip()]
        if len(row) == len(headers):
            data.append(row)
    return pd.DataFrame(data, columns=headers)

def generate_chart(df, title, fig_no):
    """根据图号自动匹配不同样式的图表"""
    plt.figure(figsize=(10, 6))
    img_stream = io.BytesIO()
    
    try:
        # 1. 分组柱状图 (图2, 图12)
        if fig_no in [2, 12]:
            df.set_index(df.columns[0]).plot(kind='bar', color=GROUPED_COLORS, ax=plt.gca(), width=0.8)
            plt.legend(bbox_to_anchor=(0.5, -0.15), loc='upper center', ncol=3, frameon=False)
        
        # 2. 饼图 (图4, 图5)
        elif fig_no in [4, 5]:
            plt.pie(df.iloc[:, 1], labels=df.iloc[:, 0], autopct='%1.1f%%', colors=plt.cm.Pastel1.colors)
        
        # 3. 柱状图 + 折线图双轴 (图9)
        elif fig_no == 9:
            fig, ax1 = plt.subplots(figsize=(10, 6))
            x = df.iloc[:, 0].astype(str)
            ax1.bar(x, df.iloc[:, 1], color=BLUE_COLOR, label='项目数量')
            ax1.set_ylabel('项目数量')
            ax2 = ax1.twinx()
            ax2.plot(x, df.iloc[:, 2], color='#ED7D31', marker='o', linewidth=2, label='到账经费')
            ax2.set_ylabel('经费(万元)')
            fig.legend(loc="upper right", bbox_to_anchor=(0.9, 0.9))
        
        # 4. 普通蓝色柱状图 (其他)
        else:
            bars = plt.bar(df.iloc[:, 0].astype(str), df.iloc[:, 1], color=BLUE_COLOR)
            for bar in bars:
                plt.text(bar.get_x() + bar.get_width()/2, bar.get_height(), 
                         f'{bar.get_height()}', ha='center', va='bottom')
        
        plt.title(title, fontsize=14, pad=20)
        plt.tight_layout()
        plt.savefig(img_stream, format='png', dpi=300)
    finally:
        plt.close()
    
    img_stream.seek(0)
    return img_stream

# =================================================================
# Word 处理逻辑
# =================================================================

def set_style(obj):
    """设置宋体小四"""
    if hasattr(obj, 'runs'):
        for run in obj.runs:
            run.font.size = Pt(12)
            run.font.name = '宋体'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

def process_content(doc, full_text):
    # 使用正则表达式拆分：标题、图名、表名、表格、普通文本
    parts = re.split(r'(\*\*图 \d+：.*?\*\*|\*\*表 \d+：.*?\*\*|## .*?|### .*?)', full_text)
    current_fig_title = None

    for part in parts:
        part = part.strip()
        if not part: continue

        if part.startswith('##'): # 二级标题
            h = doc.add_heading(part.replace('#','').strip(), level=2)
            set_style(h)
        elif part.startswith('###'): # 三级标题
            h = doc.add_heading(part.replace('#','').strip(), level=3)
            set_style(h)
        elif "**图" in part:
            current_fig_title = part.strip("*")
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(current_fig_title)
            run.bold = True
            set_style(p)
        elif "**表" in part:
            current_fig_title = None
            p = doc.add_paragraph()
            run = p.add_run(part.strip("*"))
            run.bold = True
            set_style(p)
        elif part.startswith('|'):
            df = md_table_to_df(part)
            if df is None: continue
            
            if current_fig_title: # 绘图模式
                fig_no_match = re.search(r'图 (\d+)', current_fig_title)
                fig_no = int(fig_no_match.group(1)) if fig_no_match else 0
                img = generate_chart(df, current_fig_title, fig_no)
                doc.add_picture(img, width=Inches(5.8))
                doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
                current_fig_title = None # 绘图后重置
            else: # 普通表格模式
                table = doc.add_table(rows=1, cols=len(df.columns))
                table.style = 'Table Grid'
                for i, col in enumerate(df.columns):
                    table.rows[0].cells[i].text = str(col)
                for _, row in df.iterrows():
                    row_cells = table.add_row().cells
                    for i, val in enumerate(row):
                        row_cells[i].text = str(val)
                # 设置表格内字体
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            set_style(paragraph)
        else:
            # 过滤掉一些大模型可能带出的干扰字符
            if "import matplotlib" in part or "![" in part: continue
            p = doc.add_paragraph(part.replace('$', ''))
            set_style(p)

# =================================================================
# API 接口区
# =================================================================

@app.post("/generate_report_word")
async def generate_report_word(input_data: ReportInput):
    """
    接收各章节文本，生成并返回 Word 文档流
    """
    try:
        doc = Document()
        
        # 按照 2-7 章的顺序逻辑性合并
        chapters = [
            input_data.ch2_text,
            input_data.ch3_text,
            input_data.ch4_text,
            input_data.ch5_text,
            input_data.ch6_text,
            input_data.ch7_text
        ]
        
        for content in chapters:
            if content and len(content.strip()) > 0:
                process_content(doc, content)
        
        # 将文件存入内存流
        file_stream = io.BytesIO()
        doc.save(file_stream)
        file_stream.seek(0)
        
        # 返回文件流供下载
        return StreamingResponse(
            file_stream,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={"Content-Disposition": "attachment; filename=Analysis_Report.docx"}
        )

    except Exception as e:
        error_msg = traceback.format_exc()
        print(error_msg)
        return {"status": "error", "message": str(e), "traceback": error_msg}

# --- 启动 ---
if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8080))
    uvicorn.run(app, host="0.0.0.0", port=port)
