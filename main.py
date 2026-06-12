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

# ================= 字体设置 =================
font_path = os.path.join(os.getcwd(), 'SimHei.ttf')
if os.path.exists(font_path):
    fm.fontManager.addfont(font_path)
    plt.rcParams['font.sans-serif'] = ['SimHei']
else:
    plt.rcParams['font.sans-serif'] = ['SimHei', 'Arial Unicode MS', 'DejaVu Sans']
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
    thick_sz = '12' # 1.5磅
    thin_sz = '4'   # 0.5磅
    color_val = '000000' 

    for i, row in enumerate(table.rows):
        for cell in row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            # 移除所有边框
            for b in ['top', 'bottom', 'left', 'right', 'insideH', 'insideV']:
                tag = f'w:{b}'; element = tcPr.find(qn(tag))
                if element is not None: tcPr.remove(element)
            
            # 顶线
            if i == 0:
                t = OxmlElement('w:top')
                t.set(qn('w:val'), 'single'); t.set(qn('w:sz'), thick_sz); t.set(qn('w:color'), color_val)
                tcPr.append(t)
                b = OxmlElement('w:bottom')
                b.set(qn('w:val'), 'single'); b.set(qn('w:sz'), thin_sz); b.set(qn('w:color'), color_val)
                tcPr.append(b)
            # 底线
            elif i == len(table.rows) - 1:
                b = OxmlElement('w:bottom')
                b.set(qn('w:val'), 'single'); b.set(qn('w:sz'), thick_sz); b.set(qn('w:color'), color_val)
                tcPr.append(b)

def draw_custom_pie(ax, values, labels):
    """饼图优化：解决A/B级重叠，去掉内部标签"""
    colors = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47']
    wedges, _ = ax.pie(values, colors=colors[:len(values)], startangle=90, counterclock=False, 
                       wedgeprops={'edgecolor': 'white', 'linewidth': 1.5})
    
    total = sum(values)
    for i, p in enumerate(wedges):
        ang = (p.theta2 - p.theta1)/2. + p.theta1
        y = np.sin(np.deg2rad(ang))
        x = np.cos(np.deg2rad(ang))
        
        # 增加延伸距离，防止重叠
        connectionstyle = f"angle,angleA=0,angleB={ang}"
        horizontalalignment = "left" if x > 0 else "right"
        
        # 对A级和B级这种可能非常接近的区域进行微调
        y_offset = 1.35 * y
        if abs(y) < 0.2: y_offset = 1.45 * y # 靠近水平线的进一步拉开
        
        label_text = f"{labels[i]}\n{int(values[i])} ({(values[i]/total*100):.1f}%)"
        ax.annotate(label_text, xy=(x, y), xytext=(1.5*np.sign(x), y_offset),
                    horizontalalignment=horizontalalignment,
                    arrowprops=dict(arrowstyle="-", color="black", connectionstyle=connectionstyle),
                    fontsize=10, fontweight='bold', va='center')

def process_smart(doc, text):
    # 清理公式与符号
    text = text.replace('\\%', '%').replace('$$', '').replace('\r', '')
    text = re.sub(r'\([\d\.\+\-\*/\s]+\s*[≈=]\s*[\d\.\%]+\)', '', text)
    
    lines = text.split('\n')
    table_rows, current_title, is_chart_mode = [], "", False

    def flush_table():
        nonlocal table_rows, current_title, is_chart_mode
        if not table_rows: return
        
        # 解析数据，同时过滤包含“总计”或“合计”的行（用于绘图）
        raw_data = [[c.strip() for c in r.split('|') if c.strip()] for r in table_rows if '|' in r and '---' not in r]
        
        if len(raw_data) >= 2:
            try:
                fig_num_match = re.search(r'\d+', current_title)
                fig_num = int(fig_num_match.group()) if fig_num_match else 0
                
                if is_chart_mode:
                    fig, ax = plt.subplots(figsize=(9, 5.5))
                    def clean(v):
                        s = re.sub(r'[^\d.]', '', str(v))
                        return float(s) if s else 0.0

                    # --- 图表逻辑优化 ---
                    # 1. 饼图 (图 4, 5)
                    if fig_num in [4, 5]:
                        # 过滤掉合计行
                        plot_data = [r for r in raw_data[1:] if '合计' not in r[0] and '总计' not in r[0]]
                        v_list = [clean(r[1]) for r in plot_data]
                        l_list = [r[0] for r in plot_data]
                        draw_custom_pie(ax, v_list, l_list)
                        ax.set_xlim(-2.5, 2.5); ax.set_ylim(-1.6, 1.6)

                    # 2. 图 10, 11：解决横向数据逻辑错误
                    elif fig_num in [10, 11]:
                        # 识别数据是否在第一行横向排列 (表头是年份)
                        headers = [re.sub(r'[^\d]', '', h) for h in raw_data[0]]
                        years = [h for h in headers if len(h) == 4] # 提取4位年份
                        
                        if len(years) > 1: # 说明年份在表头
                            x_labels = years
                            y_vals = [clean(v) for v in raw_data[1][1:]][:len(years)]
                        else: # 说明年份在第一列
                            x_labels = [r[0] for r in raw_data[1:]]
                            y_vals = [clean(r[1]) for r in raw_data[1:]]
                        
                        bars = ax.bar(x_labels, y_vals, color='#4472C4', width=0.5)
                        for bar in bars:
                            ax.text(bar.get_x()+bar.get_width()/2, bar.get_height(), f'{int(bar.get_height())}', ha='center', va='bottom', fontweight='bold')
                        ax.set_ylabel('数量', fontweight='bold')

                    # 3. 图 9 双轴图
                    elif fig_num == 9:
                        df = pd.DataFrame(raw_data[1:], columns=raw_data[0])
                        years = [re.sub(r'（.*?）|\(.*?\)', '', str(x)) for x in df.iloc[:, 0]]
                        counts = [clean(v) for v in df.iloc[:, 1]]
                        fundings = [clean(v) for v in df.iloc[:, 2]]
                        ax.bar(years, counts, color='#4472C4', width=0.5)
                        ax2 = ax.twinx()
                        ax2.plot(years, fundings, color='#ED7D31', marker='o', linewidth=2)
                        ax.set_ylabel('立项数 (项)')
                        ax2.set_ylabel('经费 (万元)')

                    # 4. 图 12 分组柱状图
                    elif fig_num == 12:
                        df = pd.DataFrame(raw_data[1:], columns=raw_data[0])
                        df = df[~df.iloc[:, 0].str.contains('合计|总计', na=False)]
                        x_labels = [re.sub(r'（.*?）|\(.*?\)', '', str(x)) for x in df.iloc[:, 0]]
                        categories = [c for c in df.columns[1:] if '合计' not in c]
                        x = np.arange(len(x_labels))
                        width = 0.8 / len(categories)
                        for i, cat in enumerate(categories):
                            vals = [clean(v) for v in df[cat]]
                            ax.bar(x + i*width, vals, width, label=cat)
                        ax.set_xticks(x + width*(len(categories)-1)/2)
                        ax.set_xticklabels(x_labels)
                        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.1), ncol=3)
                        plt.subplots_adjust(bottom=0.2)

                    # 统一清理：移除X轴单位
                    current_labels = [re.sub(r'（.*?）|\(.*?\)', '', str(tick.get_text())) for tick in ax.get_xticklabels()]
                    ax.set_xticklabels(current_labels)
                    
                    # 关键修改：不再 plt.title
                    plt.tight_layout()
                    buf = io.BytesIO(); plt.savefig(buf, format='png', dpi=180); plt.close(); buf.seek(0)
                    doc.add_picture(buf, width=Inches(5.8))
                    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                else:
                    # --- 表格逻辑优化 ---
                    # 过滤总计行（表 15 等要求）
                    filtered_data = [r for r in raw_data if not (len(r)>0 and ('总计' in r[0] or '合计' in r[0] and i != 0))]
                    
                    table = doc.add_table(rows=len(filtered_data), cols=len(filtered_data[0]))
                    table.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    
                    for i, row_data in enumerate(filtered_data):
                        # 检查是否是特殊行（表 10 的 B级/C级提示行）
                        row_text = "".join(row_data)
                        is_special_row = ("B级" in row_text or "C级" in row_text) and "发文数量" in row_text
                        
                        for j, val in enumerate(row_data):
                            # 去掉星号
                            clean_val = val.replace('*', '')
                            cell = table.cell(i, j)
                            cell.text = clean_val
                            p = cell.paragraphs[0]
                            
                            # 设置居中：特殊行或者标题行
                            if is_special_row or i == 0:
                                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            
                            if p.runs:
                                set_font(p.runs[0], 10, i==0 or is_special_row)
                            else:
                                set_font(p.add_run(clean_val), 10, i==0 or is_special_row)
                                
                    set_three_line_table(table)
            except Exception as e: print(f"Processing Error: {e}")
            
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

@app.post("/generate_report_word")
async def generate_report_word(input_data: ReportInput, request: Request):
    try:
        doc = Document()
        doc.settings.element.append(OxmlElement('w:updateFields'))
        
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
