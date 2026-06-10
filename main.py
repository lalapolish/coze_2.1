import re
import io
import os
import uuid
import numpy as np
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

# 环境字体设定
plt.rcParams['font.sans-serif'] = ['DejaVu Sans']
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
    """表格全边框"""
    for row in table.rows:
        for cell in row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            for b in ['top', 'bottom', 'left', 'right']:
                edge = OxmlElement(f'w:{b}')
                edge.set(qn('w:val'), 'single')
                edge.set(qn('w:sz'), '4') 
                tcPr.append(edge)

def clean_table_text(text):
    """严格清洗：只保留数字、字母、中文、百分号、点、括号"""
    if not text: return ""
    # \w 包含数字、字母、中文
    return re.sub(r'[^\w%\.\(\)]', '', str(text))

def draw_advanced_pie(ax, values, labels):
    """绘制带指引线的饼图"""
    wedges, texts = ax.pie(values, wedgeprops=dict(width=0.5), startangle=-40)
    
    bbox_props = dict(boxstyle="square,pad=0.3", fc="w", ec="k", lw=0.72)
    kw = dict(arrowprops=dict(arrowstyle="-"), bbox=bbox_props, zorder=0, va="center")

    total = sum(values)
    for i, p in enumerate(wedges):
        ang = (p.theta2 - p.theta1)/2. + p.theta1
        y = np.sin(np.deg2rad(ang))
        x = np.cos(np.deg2rad(ang))
        horizontalalignment = {-1: "right", 1: "left"}[int(np.sign(x))]
        connectionstyle = f"angle,angleA=0,angleB={ang}"
        kw["arrowprops"].update({"connectionstyle": connectionstyle})
        
        # 标注内容：类别, 数值, 百分比
        pct = values[i]/total*100
        val_str = f"{labels[i]}\n{int(values[i])}, {pct:.1f}%"
        
        ax.annotate(val_str, xy=(x, y), xytext=(1.35*np.sign(x), 1.4*y),
                    horizontalalignment=horizontalalignment, **kw)

def process_smart(doc, text):
    text = text.replace('$$', '').replace('\r', '')
    lines = text.split('\n')
    
    table_rows = []
    current_title = ""
    is_chart_mode = False

    def flush_table():
        nonlocal table_rows, current_title, is_chart_mode
        if not table_rows: return
        
        raw_data = []
        for r in table_rows:
            cells = [c.strip() for c in r.split('|') if c.strip()]
            if cells and '---' not in r: raw_data.append(cells)
        
        if len(raw_data) >= 2:
            try:
                df = pd.DataFrame(raw_data[1:], columns=raw_data[0])
                # 提取标题中的数字
                fig_num_match = re.search(r'\d+', current_title)
                fig_num = int(fig_num_match.group()) if fig_num_match else 0

                if is_chart_mode:
                    fig, ax = plt.subplots(figsize=(9, 5))
                    # 转换数值
                    for col in df.columns[1:]:
                        df[col] = pd.to_numeric(df[col].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
                    
                    # 判定逻辑
                    is_pie = ("分布" in current_title or "比例" in current_title) and fig_num not in [3, 6]
                    
                    if is_pie:
                        # 饼图模式
                        draw_advanced_pie(ax, df.iloc[:, 1].values, df.iloc[:, 0].values)
                        ax.set_xlim(-2, 2)
                        ax.set_ylim(-1.5, 1.5)
                    else:
                        # 柱状图模式
                        x_indices = np.arange(len(df))
                        if len(df.columns) > 2:
                            # 分组柱状图 (图 2, 12 等)
                            width = 0.8 / (len(df.columns) - 1)
                            for i, col in enumerate(df.columns[1:]):
                                ax.bar(x_indices + i*width, df[col], width, label=col)
                            ax.set_xticks(x_indices + width*(len(df.columns)-2)/2)
                            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.1), ncol=5)
                        else:
                            # 普通柱状图
                            ax.bar(df.iloc[:, 0].astype(str), df.iloc[:, 1], color='#4472C4', width=0.5)
                            for i, val in enumerate(df.iloc[:, 1]):
                                ax.text(i, val, f'{int(val)}', ha='center', va='bottom')

                    plt.tight_layout()
                    buf = io.BytesIO()
                    plt.savefig(buf, format='png', dpi=180)
                    plt.close()
                    buf.seek(0)
                    doc.add_picture(buf, width=Inches(5.8))
                    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
                else:
                    # 表格清洗与生成
                    table = doc.add_table(rows=len(raw_data), cols=len(raw_data[0]))
                    table.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for i, row_data in enumerate(raw_data):
                        for j, val in enumerate(row_data):
                            cell = table.cell(i, j)
                            # 严格清洗内容
                            cell.text = clean_table_text(val)
                            set_font(cell.paragraphs[0].runs[0], 10, i==0)
                    set_full_border(table)
            except Exception as e:
                print(f"Error: {e}")
        
        table_rows, is_chart_mode, current_title = [], False, ""

    for line in lines:
        l = line.strip()
        if not l: continue
        if l.startswith('#'):
            flush_table()
            p = doc.add_heading('', level=min(l.count('#'), 3))
            run = p.add_run(l.replace('#', '').strip())
            set_font(run, 14, True)
        elif re.match(r'(\*\*?)?[图表]\s?\d+[:：]', l):
            flush_table()
            current_title = l.replace('*', '').strip()
            is_chart_mode = "图" in current_title
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(current_title)
            set_font(run, 11, True)
        elif l.startswith('|'):
            table_rows.append(l)
        else:
            if table_rows: flush_table()
            p = doc.add_paragraph()
            p.paragraph_format.first_line_indent = Pt(24)
            run = p.add_run(l.replace('**', ''))
            set_font(run, 12)

    flush_table()

@app.post("/generate_report_word")
async def generate_report_word(input_data: ReportInput, request: Request):
    try:
        doc = Document()
        # 逐章节处理
        has_content = False
        for i in range(2, 8):
            txt = getattr(input_data, f"ch{i}_text", "")
            if txt and txt.strip():
                process_smart(doc, txt)
                has_content = True
        
        if not has_content:
            return {"file": "", "message": "输入内容为空", "status": "error"}

        # 生成随机文件名
        fname = f"report_{uuid.uuid4().hex[:8]}.docx"
        static_dir = "static"
        if not os.path.exists(static_dir):
            os.makedirs(static_dir)
            
        full_path = os.path.join(static_dir, fname)
        doc.save(full_path)
        
        # 拼接完整的下载 URL
        # 这里的 key 必须叫做 file，才能对应你在工作流里定义的输出变量名
        file_url = f"{str(request.base_url).rstrip('/')}/static/{fname}"
        
        return {
            "file": file_url,  # 这里的键名必须与插件面板定义的输出参数名一致
            "filename": fname,
            "status": "success"
        }
    except Exception as e:
        # 即使报错也建议返回一个空的 file 字段防止 workflow 彻底卡死
        return {"file": "", "status": "error", "message": str(e)}


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8080)
