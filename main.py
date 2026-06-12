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
from docx.shared import Pt, Inches
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
    """设置标准的黑色加粗三线表"""
    thick_sz = '18' # 2.25pt
    thin_sz = '8'   # 1pt
    color_val = '000000'
    for i, row in enumerate(table.rows):
        for cell in row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            for b in ['top', 'bottom', 'left', 'right', 'insideH', 'insideV']:
                tag = f'w:{b}'; element = tcPr.find(qn(tag))
                if element is not None: tcPr.remove(element)
            if i == 0:
                for pos, sz in [('top', thick_sz), ('bottom', thin_sz)]:
                    border = OxmlElement(f'w:{pos}')
                    border.set(qn('w:val'), 'single'); border.set(qn('w:sz'), sz); border.set(qn('w:color'), color_val)
                    tcPr.append(border)
            if i == len(table.rows) - 1:
                btm = OxmlElement('w:bottom')
                btm.set(qn('w:val'), 'single'); btm.set(qn('w:sz'), thick_sz); btm.set(qn('w:color'), color_val)
                tcPr.append(btm)

def draw_custom_pie(ax, values, labels):
    """饼图优化：数值延伸、加粗、解决遮挡"""
    colors = ['#4472C4', '#ED7D31', '#FFD966', '#E8307E', '#2EBB9F', '#2E3192']
    wedges, _ = ax.pie(values, colors=colors[:len(values)], startangle=90, counterclock=False, 
                       wedgeprops={'edgecolor': 'white', 'linewidth': 1.5})
    
    total = sum(values)
    for i, p in enumerate(wedges):
        ang = (p.theta2 - p.theta1)/2. + p.theta1
        y = np.sin(np.deg2rad(ang))
        x = np.cos(np.deg2rad(ang))
        horizontalalignment = {-1: "right", 1: "left"}[int(np.sign(x))]
        connectionstyle = f"angle,angleA=0,angleB={ang}"
        label_text = f"{labels[i]}\n{int(values[i])} ({(values[i]/total*100):.1f}%)"
        ax.annotate(label_text, xy=(x, y), xytext=(1.35*np.sign(x), 1.35*y),
                    horizontalalignment=horizontalalignment,
                    arrowprops=dict(arrowstyle="-", color="black", connectionstyle=connectionstyle),
                    fontsize=11, fontweight='bold', va='center')

def add_page_number(doc):
    """添加页码功能"""
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
                    fig, ax = plt.subplots(figsize=(11, 7))
                    def clean(v):
                        s = re.sub(r'[^\d.]', '', str(v))
                        return float(s) if s else 0.0

                    # 逻辑 1：图 4, 5 饼图
                    if fig_num in [4, 5]:
                        v_list = [clean(v) for v in df.iloc[:, 1]]
                        draw_custom_pie(ax, v_list, df.iloc[:, 0].tolist())
                        ax.set_xlim(-2.2, 2.2); ax.set_ylim(-1.8, 1.8)

                    # 逻辑 2：图 9 双轴图 (立项数柱状 + 经费折线)
                    elif fig_num == 9:
                        years = df.iloc[:, 0].astype(str).tolist()
                        counts = [clean(v) for v in df.iloc[:, 1]]
                        fundings = [clean(v) for v in df.iloc[:, 2]]
                        
                        ax.bar(years, counts, color='#4472C4', label='立项数(项)', width=0.5)
                        ax.set_xlabel('年份(publish_year)', fontsize=12, fontweight='bold')
                        ax.set_ylabel('立项数(项)(count)', fontsize=12, fontweight='bold')
                        for i, v in enumerate(counts):
                            ax.text(i, v, f'{int(v)}', ha='center', va='bottom', fontweight='bold')
                        
                        ax2 = ax.twinx()
                        ax2.plot(years, fundings, color='#ED7D31', marker='o', linewidth=2, label='到账经费(万元)')
                        ax2.set_ylabel('到账经费(万元)(received_funding)', fontsize=12, fontweight='bold')
                        for i, v in enumerate(fundings):
                            ax2.text(i, v, f'{v:.1f}', ha='center', va='bottom', color='#ED7D31', fontweight='bold')
                        fig.legend(loc='upper center', bbox_to_anchor=(0.5, 0.98), ncol=2)

                    # 逻辑 3：其他柱状图（含图 12）
                    else:
                        cols = [c for c in df.columns if '合计' not in c and '总计' not in c]
                        df_plot = df[cols][~df.iloc[:, 0].str.contains('合计|总计', na=False)]
                        x_labels = df_plot.iloc[:, 0].astype(str).tolist()
                        
                        if len(df_plot.columns) > 2: # 多系列（图 12）
                            x = np.arange(len(df_plot))
                            width = 0.8 / (len(df_plot.columns)-1)
                            for i, col in enumerate(df_plot.columns[1:]):
                                vals = [clean(v) for v in df_plot[col]]
                                rects = ax.bar(x + i*width, vals, width, label=col)
                                for rect in rects:
                                    h = rect.get_height()
                                    if h > 0: ax.text(rect.get_x()+rect.get_width()/2, h, f'{int(h)}', ha='center', va='bottom', fontsize=9, fontweight='bold')
                            ax.set_xticks(x + width*(len(df_plot.columns)-2)/2)
                            ax.set_xticklabels(x_labels)
                            ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)
                        else: # 单系列
                            vals = [clean(v) for v in df_plot.iloc[:, 1]]
                            bars = ax.bar(x_labels, vals, color='#4472C4', width=0.5)
                            for bar in bars:
                                ax.text(bar.get_x()+bar.get_width()/2, bar.get_height(), f'{int(bar.get_height())}', ha='center', va='bottom', fontweight='bold')
                        
                        ax.set_xlabel(df_plot.columns[0], fontsize=12, fontweight='bold')
                        ax.set_ylabel(df_plot.columns[1] if len(df_plot.columns)>1 else "数量", fontsize=12, fontweight='bold')

                    plt.title(current_title, fontsize=14, fontweight='bold', pad=20)
                    plt.tight_layout()
                    buf = io.BytesIO(); plt.savefig(buf, format='png', dpi=180); plt.close(); buf.seek(0)
                    doc.add_picture(buf, width=Inches(5.8))
                    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
                else:
                    table = doc.add_table(rows=len(raw_data), cols=len(raw_data[0]))
                    table.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for i, row in enumerate(raw_data):
                        for j, val in enumerate(row):
                            cell = table.cell(i, j); cell.text = val
                            set_font(cell.paragraphs[0].runs[0] if cell.paragraphs[0].runs else cell.paragraphs[0].add_run(val), 10, i==0)
                    set_three_line_table(table)
            except Exception as e: print(f"Plot Error: {e}")
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
