import re
import io
import os
import uuid
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
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
    ch1_text: Optional[str] = ""  
    ch2_text: Optional[str] = ""
    ch3_text: Optional[str] = ""
    ch4_text: Optional[str] = ""
    ch5_text: Optional[str] = ""
    ch6_text: Optional[str] = ""
    ch7_text: Optional[str] = ""
    ch8_text: Optional[str] = "" 

# ================= 解决图片方框：强制加载本地 SimHei.ttf =================
font_path = os.path.join(os.getcwd(), 'SimHei.ttf')
if os.path.exists(font_path):
    fm.fontManager.addfont(font_path)
    plt.rcParams['font.sans-serif'] = ['SimHei']
else:
    plt.rcParams['font.sans-serif'] = ['SimHei', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

def set_font(run, size=12, bold=False):
    run.font.size = Pt(size)
    run.font.name = '宋体'
    run.bold = bold
    rPr = run._element.get_or_add_rPr()
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:eastAsia'), '宋体')
    rPr.append(rFonts)

def set_three_line_table(table):
    """设置学术标准的黑色三线表 (顶线/底线加粗，标题下线细)"""
    def set_cell_border(cell, **kwargs):
        tcPr = cell._tc.get_or_add_tcPr()
        for side, params in kwargs.items():
            tag = f'w:{side}'
            element = OxmlElement(tag)
            element.set(qn('w:val'), params.get('val', 'single'))
            element.set(qn('w:sz'), str(params.get('sz', '4')))
            element.set(qn('w:color'), params.get('color', '000000'))
            tcPr.append(element)

    for i, row in enumerate(table.rows):
        for cell in row.cells:
            # 清除所有默认边框
            tcPr = cell._tc.get_or_add_tcPr()
            for b in ['top', 'bottom', 'left', 'right', 'insideH', 'insideV']:
                tag = f'w:{b}'; element = tcPr.find(qn(tag))
                if element is not None: tcPr.remove(element)
            
            # 顶层加粗 (1.5pt = sz:12)
            if i == 0:
                set_cell_border(cell, top={'sz': 12, 'color': '000000'}, bottom={'sz': 6, 'color': '000000'})
            # 底层加粗 (1.5pt = sz:12)
            elif i == len(table.rows) - 1:
                set_cell_border(cell, bottom={'sz': 12, 'color': '000000'})

def draw_custom_pie(ax, values, labels):
    """饼图：彻底解决数值重叠和方框问题"""
    colors = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47']
    wedges, _ = ax.pie(values, colors=colors[:len(values)], startangle=90, counterclock=False, 
                       wedgeprops={'edgecolor': 'white', 'linewidth': 1})
    
    total = sum(values)
    for i, p in enumerate(wedges):
        ang = (p.theta2 - p.theta1)/2. + p.theta1
        y = np.sin(np.deg2rad(ang))
        x = np.cos(np.deg2rad(ang))
        
        # 动态计算标签位置，防止靠得太近
        va = "center"; ha = "left" if x > 0 else "right"
        dist = 1.4  # 延伸长度
        connectionstyle = f"angle,angleA=0,angleB={ang}"
        
        label_text = f"{labels[i]}\n{int(values[i])} ({(values[i]/total*100):.1f}%)"
        ax.annotate(label_text, xy=(x, y), xytext=(dist*np.sign(x), 1.3*y),
                    horizontalalignment=ha, verticalalignment=va,
                    arrowprops=dict(arrowstyle="-", color="black", connectionstyle=connectionstyle),
                    fontsize=10, fontweight='bold')

def add_page_number(doc):
    for sec in doc.sections:
        footer = sec.footer
        p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run()
        fldChar1 = OxmlElement('w:fldChar'); fldChar1.set(qn('w:fldCharType'), 'begin')
        instrText = OxmlElement('w:instrText'); instrText.set(qn('xml:space'), 'preserve'); instrText.text = 'PAGE'
        fldChar2 = OxmlElement('w:fldChar'); fldChar2.set(qn('w:fldCharType'), 'end')
        run._r.extend([fldChar1, instrText, fldChar2])

def process_smart(doc, text):
    text = text.replace('\\%', '%').replace('$$', '').replace('\r', '')
    text = re.sub(r'\([\d\.\+\-\*/\s]+\s*[≈=]\s*[\d\.\%]+\)', '', text)
    
    lines = text.split('\n')
    table_rows, current_title, is_chart_mode = [], "", False

    def flush_table():
        nonlocal table_rows, current_title, is_chart_mode
        if not table_rows: return
        raw_data = [[c.strip() for c in r.split('|') if c.strip()] for r in table_rows if '|' in r and '---' not in r]
        
        if len(raw_data) >= 2:
            try:
                df = pd.DataFrame(raw_data[1:], columns=raw_data[0])
                fig_num_match = re.search(r'\d+', current_title)
                fig_num = int(fig_num_match.group()) if fig_num_match else 0
                
                if is_chart_mode:
                    fig, ax = plt.subplots(figsize=(10, 6))
                    def clean(v):
                        s = re.sub(r'[^\d.]', '', str(v))
                        return float(s) if s else 0.0

                    # 1. 饼图 (4, 5)
                    if fig_num in [4, 5]:
                        v_list = [clean(v) for v in df.iloc[:, 1]]
                        draw_custom_pie(ax, v_list, df.iloc[:, 0].tolist())
                        ax.set_xlim(-2.2, 2.2); ax.set_ylim(-1.5, 1.5)

                    # 2. 图 9 双轴图：优化重叠
                    elif fig_num == 9:
                        years = df.iloc[:, 0].astype(str).tolist()
                        counts = [clean(v) for v in df.iloc[:, 1]]
                        fundings = [clean(v) for v in df.iloc[:, 2]]
                        ax.bar(years, counts, color='#4472C4', label='立项数(项)', width=0.5)
                        ax.set_ylabel('立项数 (项)', fontsize=11, fontweight='bold')
                        for i, v in enumerate(counts):
                            ax.text(i, v + 2, f'{int(v)}', ha='center', va='bottom', fontweight='bold', color='#4472C4')
                        
                        ax2 = ax.twinx()
                        ax2.plot(years, fundings, color='#ED7D31', marker='o', linewidth=2, label='到账经费(万元)')
                        ax2.set_ylabel('经费 (万元)', fontsize=11, fontweight='bold')
                        for i, v in enumerate(fundings):
                            ax2.text(i, v + 50, f'{v:.1f}', ha='center', va='bottom', color='#ED7D31', fontweight='bold')
                        
                        ax.set_ylim(0, max(counts)*1.3)
                        ax2.set_ylim(0, max(fundings)*1.3)
                        fig.legend(loc='lower center', bbox_to_anchor=(0.5, 0.02), ncol=2)
                        plt.subplots_adjust(top=0.88, bottom=0.15)

                    # 3. 图 10, 11：强制单系列柱状图
                    elif fig_num in [10, 11]:
                        # 只取年份和对应的总量（通常是第1或最后一列）
                        x_labels = df.iloc[:, 0].astype(str).tolist()
                        y_vals = [clean(v) for v in df.iloc[:, 1]]
                        bars = ax.bar(x_labels, y_vals, color='#4472C4', width=0.5)
                        for bar in bars:
                            ax.text(bar.get_x()+bar.get_width()/2, bar.get_height(), f'{int(bar.get_height())}', ha='center', va='bottom', fontweight='bold')
                        ax.set_ylabel('数量', fontweight='bold')
                        ax.set_xlabel('年份', fontweight='bold')

                    # 4. 图 12：多系列分组柱状图 (A-F级)
                    elif fig_num == 12:
                        # 过滤非数据列，保留年份和各级列
                        df_plot = df[~df.iloc[:, 0].str.contains('合计|总计', na=False)]
                        x_labels = df_plot.iloc[:, 0].astype(str).tolist()
                        categories = [c for c in df_plot.columns[1:] if '合计' not in c]
                        
                        x = np.arange(len(x_labels))
                        width = 0.75 / len(categories)
                        for i, cat in enumerate(categories):
                            vals = [clean(v) for v in df_plot[cat]]
                            rects = ax.bar(x + i*width, vals, width, label=cat)
                            for rect in rects:
                                h = rect.get_height()
                                if h > 0: ax.text(rect.get_x()+rect.get_width()/2, h, f'{int(h)}', ha='center', va='bottom', fontsize=8)
                        
                        ax.set_xticks(x + width*(len(categories)-1)/2)
                        ax.set_xticklabels(x_labels)
                        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.12), ncol=len(categories))
                        ax.set_ylabel('获奖数量', fontweight='bold')
                        plt.subplots_adjust(bottom=0.2)

                    # 其他通用图表
                    else:
                        vals = [clean(v) for v in df.iloc[:, 1]]
                        ax.bar(df.iloc[:, 0].astype(str), vals, color='#4472C4', width=0.5)
                        for i, v in enumerate(vals): ax.text(i, v, f'{int(v)}', ha='center', va='bottom')

                    plt.title(current_title, fontsize=12, fontweight='bold', pad=15)
                    buf = io.BytesIO(); plt.savefig(buf, format='png', dpi=150, bbox_inches='tight'); plt.close(); buf.seek(0)
                    doc.add_picture(buf, width=Inches(5.6))
                    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
                else:
                    table = doc.add_table(rows=len(raw_data), cols=len(raw_data[0]))
                    table.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for i, row in enumerate(raw_data):
                        for j, val in enumerate(row):
                            cell = table.cell(i, j); cell.text = val
                            set_font(cell.paragraphs[0].runs[0] if cell.paragraphs[0].runs else cell.paragraphs[0].add_run(val), 10, i==0)
                    set_three_line_table(table)
            except Exception as e: print(f"Error processing {current_title}: {e}")
        table_rows, is_chart_mode, current_title = [], False, ""

    for line in lines:
        l = line.strip()
        if not l: continue
        if l.startswith('#'):
            flush_table(); p = doc.add_heading('', level=min(l.count('#'), 3))
            run = p.add_run(l.replace('#', '').strip()); set_font(run, 14, True)
        elif re.match(r'(\*\*?)?[图表]\s?\d+[:：\s]', l):
            flush_table(); current_title = l.replace('*', '').strip(); is_chart_mode = "图" in current_title
            p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(current_title); set_font(run, 11, True)
        elif l.startswith('|'): table_rows.append(l)
        else:
            if table_rows: flush_table()
            p = doc.add_paragraph(); p.paragraph_format.first_line_indent = Pt(24)
            run = p.add_run(l.replace('**', '')); set_font(run, 12)
    flush_table()

def add_toc(doc):
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("目录"); set_font(run, size=16, bold=True)
    p_toc = doc.add_paragraph(); run_toc = p_toc.add_run()
    fldChar1 = OxmlElement('w:fldChar'); fldChar1.set(qn('w:fldCharType'), 'begin')
    instrText = OxmlElement('w:instrText'); instrText.set(qn('xml:space'), 'preserve'); instrText.text = 'TOC \\o "1-3" \\h \\z \\u'
    fldChar2 = OxmlElement('w:fldChar'); fldChar2.set(qn('w:fldCharType'), 'separate')
    fldChar3 = OxmlElement('w:fldChar'); fldChar3.set(qn('w:fldCharType'), 'end')
    run_toc._r.extend([fldChar1, instrText, fldChar2, fldChar3])
    doc.add_page_break()

@app.post("/generate_report_word")
async def generate_report_word(input_data: ReportInput, request: Request):
    try:
        doc = Document()
        element = doc.settings.element
        update_fields = OxmlElement('w:updateFields'); update_fields.set(qn('w:val'), 'true'); element.append(update_fields)
        
        # 封面
        for _ in range(4): doc.add_paragraph()
        p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run("山东师范大学人文社会科学科研成果发展态势分析报告\n（2020-2024）")
        set_font(run, size=22, bold=True); doc.add_page_break()
        
        add_toc(doc)
        add_page_number(doc)
        
        for i in range(1, 9): 
            txt = getattr(input_data, f"ch{i}_text", "")
            if txt and txt.strip(): process_smart(doc, txt)
        
        fname = f"report_{uuid.uuid4().hex[:8]}.docx"
        full_path = os.path.join("static", fname)
        doc.save(full_path)
        return {"file": f"{str(request.base_url).rstrip('/')}/static/{fname}", "status": "success"}
    except Exception as e: return {"file": "", "status": "error", "message": str(e)}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8080)
