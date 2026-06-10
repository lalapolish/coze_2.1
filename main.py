import re
import io
import os
import pandas as pd
import matplotlib.pyplot as plt
from fastapi import FastAPI, HTTPException
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
from typing import Optional, List
from docx import Document
from docx.shared import Pt, Inches
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH

# --- 初始化 FastAPI (使用 3.0.0 兼容扣子) ---
app = FastAPI(openapi_version="3.0.0")

# --- 输入模型定义 ---
class ReportInput(BaseModel):
    ch2_text: Optional[str] = ""
    ch3_text: Optional[str] = ""
    ch4_text: Optional[str] = ""
    ch5_text: Optional[str] = ""
    ch6_text: Optional[str] = ""
    ch7_text: Optional[str] = ""

# --- 绘图全局配置 ---
plt.rcParams['font.sans-serif'] = ['SimHei']  # 确保 Render 环境中有中文字体，否则请改为通用字体
plt.rcParams['axes.unicode_minus'] = False
BLUE_COLOR = 'royalblue'
GROUPED_COLORS = ['#4472C4', '#ED7D31', '#FFC000', '#70AD47', '#25B6C7']

def clean_data(text):
    """清洗 Markdown 表格中的单元格数据"""
    if pd.isna(text): return 0
    # 移除可能干扰转化的符号，如 $ % , 
    clean_val = str(text).replace('$', '').replace(',', '').replace('%', '').strip()
    try:
        return float(clean_val) if '.' in clean_val else int(clean_val)
    except:
        return clean_val

def md_table_to_df(md_text):
    """将 Markdown 格式表格转换为 Pandas DataFrame"""
    lines = [line.strip() for line in md_text.strip().split('\n') if '|' in line]
    if len(lines) < 2: return None
    # 提取表头并去除多余字符
    headers = [re.sub(r'[\$\*]', '', c).strip() for c in lines[0].split('|') if c.strip()]
    data = []
    # 跳过表头和分隔行
    for line in lines[2:]:
        row = [clean_data(c) for c in line.split('|') if c.strip()]
        if len(row) == len(headers):
            data.append(row)
    return pd.DataFrame(data, columns=headers)

def generate_chart(df, title, fig_no):
    """根据 DataFrame 生成图表并返回二进制流"""
    plt.figure(figsize=(10, 6))
    img_stream = io.BytesIO()
    try:
        # 逻辑判断：根据图号决定绘图类型（示例：4/5号图画饼图，其他画柱状图）
        if fig_no in [2, 12]:
            df.set_index(df.columns[0]).plot(kind='bar', color=GROUPED_COLORS, ax=plt.gca(), width=0.8)
            plt.legend(bbox_to_anchor=(0.5, -0.15), loc='upper center', ncol=3, frameon=False)
        elif fig_no in [4, 5]:
            plt.pie(df.iloc[:, 1], labels=df.iloc[:, 0], autopct='%1.1f%%', colors=plt.cm.Pastel1.colors)
        else:
            # 默认：蓝色柱状图
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
    """统一设置 Word 文档字体为宋体 12pt (小四)"""
    if hasattr(obj, 'runs'):
        for run in obj.runs:
            run.font.size = Pt(12)
            run.font.name = '宋体'
            # 强制设置中文字体
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

def process_content(doc, full_text):
    """核心：解析文本内容，识别标题、文本、表格并自动转化为图表"""
    # 按照标题、图表标记、普通文本进行切割
    parts = re.split(r'(\*\*图 \d+：.*?\*\*|\*\*表 \d+：.*?\*\*|## .*?|### .*?)', full_text)
    current_fig_title = None
    
    for part in parts:
        part = part.strip()
        if not part: continue
        
        if part.startswith('##'):
            # 处理标题
            h = doc.add_heading(part.replace('#','').strip(), level=2)
            set_style(h)
        elif "**图" in part:
            # 识别图表标题并记录，准备给接下来的表格绘图
            current_fig_title = part.strip("*")
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(current_fig_title)
            run.bold = True
            set_style(p)
        elif part.startswith('|'):
            # 处理表格
            df = md_table_to_df(part)
            if df is not None:
                if current_fig_title:
                    # 如果上方有图表标题，则生成图片
                    fig_no_match = re.search(r'图 (\d+)', current_fig_title)
                    fig_no = int(fig_no_match.group(1)) if fig_no_match else 0
                    img = generate_chart(df, current_fig_title, fig_no)
                    doc.add_picture(img, width=Inches(5.8))
                    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    current_fig_title = None # 绘图完清空标记
                else:
                    # 否则生成普通 Word 表格
                    table = doc.add_table(rows=1, cols=len(df.columns))
                    table.style = 'Table Grid'
                    for i, col in enumerate(df.columns): 
                        table.rows[0].cells[i].text = str(col)
                    for _, row in df.iterrows():
                        row_cells = table.add_row().cells
                        for i, val in enumerate(row): 
                            row_cells[i].text = str(val)
        else:
            # 处理普通段落
            p = doc.add_paragraph(part.replace('$', ''))
            set_style(p)

@app.post("/generate_report_word")
async def generate_report_word(input_data: ReportInput):
    """主接口：接收各章节文本，生成并返回 Word 文件流"""
    try:
        doc = Document()
        # 按顺序组合各章节
        chapters = [
            input_data.ch2_text, input_data.ch3_text, 
            input_data.ch4_text, input_data.ch5_text, 
            input_data.ch6_text, input_data.ch7_text
        ]
        
        for content in chapters:
            if content and len(content.strip()) > 0:
                process_content(doc, content)
        
        # 将文档保存到内存流
        file_stream = io.BytesIO()
        doc.save(file_stream)
        file_stream.seek(0)
        
        return StreamingResponse(
            file_stream, 
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document", 
            headers={"Content-Disposition": "attachment; filename=Analysis_Report.docx"}
        )
    except Exception as e:
        return {"status": "error", "message": str(e)}

if __name__ == "__main__":
    import uvicorn
    # 端口默认为 8080，适配 Render
    uvicorn.run(app, host="0.0.0.0", port=int(os.environ.get("PORT", 8080)))
