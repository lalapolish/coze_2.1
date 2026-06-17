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
from docx.shared import Pt, Inches, RGBColor, Cm
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL

app = FastAPI(openapi_version="3.0.0")

# 确保静态资源目录存在
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
    appendix: Optional[str] = "" 

# ================= 字体与全局配置 =================
font_path = os.path.join(os.getcwd(), 'SimHei.ttf')
if os.path.exists(font_path):
    fm.fontManager.addfont(font_path)
    plt.rcParams['font.sans-serif'] = ['SimHei']
else:
    plt.rcParams['font.sans-serif'] = ['SimHei', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

def set_font(run, size=12, bold=False, color="000000"):
    """设置字体样式：强制黑色，确保中西文一致"""
    run.font.size = Pt(size)
    run.font.name = '宋体'
    run.bold = bold
    # 强制所有字体颜色为黑色
    run.font.color.rgb = RGBColor.from_string("000000")
    
    rPr = run._element.get_or_add_rPr()
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:eastAsia'), '宋体')
    rPr.append(rFonts)

def set_cell_shading(cell, color):
    """设置单元格背景颜色（底纹）"""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color)
    tcPr.append(shd)

def set_table_border(table):
    """设置表格蓝色边框线"""
    tbl = table._tbl
    ptr = tbl.get_or_add_tblPr()
    borders = OxmlElement('w:tblBorders')
    
    # 外边框：蓝色 1.5pt
    for border_name in ['top', 'bottom', 'left', 'right']:
        edge = OxmlElement(f'w:{border_name}')
        edge.set(qn('w:val'), 'single')
        edge.set(qn('w:sz'), '12') 
        edge.set(qn('w:color'), '4472C4')
        borders.append(edge)
    
    # 内部横线：蓝色 0.5pt
    inside_h = OxmlElement('w:insideH')
    inside_h.set(qn('w:val'), 'single')
    inside_h.set(qn('w:sz'), '4') 
    inside_h.set(qn('w:color'), '4472C4')
    borders.append(inside_h)
    
    ptr.append(borders)

def draw_custom_pie(ax, values, labels, fig_num=0):
    """饼图绘制：带连接线的扇形图"""
    colors = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47']
    wedges, _ = ax.pie(values, colors=colors[:len(values)], startangle=90, counterclock=False, 
                       wedgeprops={'edgecolor': 'white', 'linewidth': 1})
    
    total = sum(values)
    for i, p in enumerate(wedges):
        ang = (p.theta2 - p.theta1)/2. + p.theta1
        y = np.sin(np.deg2rad(ang))
        x = np.cos(np.deg2rad(ang))
        
        dist = 1.7
        y_text = 1.35 * y
        if abs(y) < 0.3:
            y_text = 1.5 * y
        
        x_text = dist * np.sign(x) if x != 0 else dist
        ha = "left" if x_text > 0 else "right"
        
        # 针对图5 A/B 级的特殊偏移
        if fig_num == 5:
            lbl_str = str(labels[i]).strip()
            if lbl_str == 'A':
                x_text, ha, y_text = 1.7, "left", 1.6
            elif lbl_str == 'B':
                x_text, ha, y_text = -1.7, "right", 1.6

        label_text = f"{labels[i]}\n{int(values[i])} ({(values[i]/total*100):.1f}%)"
        ax.annotate(label_text, xy=(x, y), xytext=(x_text, y_text),
                    horizontalalignment=ha, verticalalignment="center",
                    arrowprops=dict(arrowstyle="-", color="black", connectionstyle=f"angle,angleA=0,angleB={ang}"),
                    fontsize=10, fontweight='bold', color='black')

def add_page_number(doc):
    """页脚插入页码"""
    for sec in doc.sections:
        footer = sec.footer
        p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run()
        f1 = OxmlElement('w:fldChar'); f1.set(qn('w:fldCharType'), 'begin')
        inst = OxmlElement('w:instrText'); inst.set(qn('xml:space'), 'preserve'); inst.text = 'PAGE'
        f2 = OxmlElement('w:fldChar'); f2.set(qn('w:fldCharType'), 'end')
        run._r.extend([f1, inst, f2])

def process_smart(doc, text):
    """核心处理函数：解析 Markdown 结构并转为 Word 对象"""
    # 预清洗：处理用户输入中可能干扰解析的标签和符号
    text = text.replace('<center>', '').replace('</center>', '')
    text = text.replace('\\%', '%').replace('$$', '').replace('\r', '')
    text = re.sub(r'\([\d\.\+\-\*/\s]+\s*[≈=]\s*[\d\.\%]+\)', '', text)
    
    lines = text.split('\n')
    table_rows = []
    current_title = ""
    is_chart_mode = False

    def flush_table():
        nonlocal table_rows, current_title, is_chart_mode
        if not table_rows:
            return
        
        # 提取表格行，并严格对齐列数防止 DataFrame 报错
        raw_data = []
        for r in table_rows:
            if '|' in r and '---' not in r:
                cells = [c.strip() for c in r.split('|') if c.strip()]
                if cells:
                    raw_data.append(cells)
        
        if len(raw_data) >= 2:
            try:
                header = raw_data[0]
                # 鲁棒性：补齐不完整的行数据
                data_cleaned = []
                for row in raw_data[1:]:
                    if len(row) > len(header):
                        data_cleaned.append(row[:len(header)])
                    else:
                        data_cleaned.append(row + [""] * (len(header) - len(row)))
                
                df = pd.DataFrame(data_cleaned, columns=header)
                num_match = re.search(r'\d+', current_title)
                num_val = int(num_match.group()) if num_match else 0
                
                if is_chart_mode:
                    fig, ax = plt.subplots(figsize=(10, 6))
                    
                    # 【核心修改】：添加背景灰色条纹网格线（置于底层）
                    if num_val not in [4, 5]:
                        ax.yaxis.grid(True, linestyle='-', which='major', color='#D9D9D9', alpha=0.6)
                        ax.set_axisbelow(True) 

                    def val_c(v):
                        s = re.sub(r'[^\d.]', '', str(v))
                        return float(s) if s else 0.0

                    # 1. 饼图分支 (图4, 5)
                    if num_val in [4, 5]:
                        v_list = [val_c(v) for v in df.iloc[:, 1]]
                        draw_custom_pie(ax, v_list, df.iloc[:, 0].tolist(), num_val)
                        ax.set_xlim(-2.5, 2.5); ax.set_ylim(-1.6, 1.6)

                    # 2. 多系列分组柱状图 (图2, 12)
                    elif num_val in [2, 12]:
                        c_list = ['#F4CCCC', '#F9CB9C', '#FFE599', '#B6D7A8', '#A2C4C9', '#A4C2F4']
                        df_plot = df[~df.iloc[:, 0].str.contains('合计|总计', na=False)]
                        x_labs = [re.sub(r'（.*?）|\(.*?\)|万元|项', '', str(x)) for x in df_plot.iloc[:, 0]]
                        cats = [c for c in df_plot.columns[1:] if '合计' not in c]
                        x_pos = np.arange(len(x_labs))
                        bar_w = 0.8 / (len(cats) + 1)
                        for idx, cat in enumerate(cats):
                            vs = [val_c(v) for v in df_plot[cat]]
                            ax.bar(x_pos + idx*bar_w, vs, bar_w, label=cat, color=c_list[idx % len(c_list)])
                        ax.set_xticks(x_pos + bar_w*(len(cats)-1)/2)
                        ax.set_xticklabels(x_labs, color='black')
                        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=len(cats))
                        plt.subplots_adjust(bottom=0.2)

                    # 3. 双轴折线柱状图 (图9)
                    elif num_val == 9:
                        yrs = [re.sub(r'（.*?）|\(.*?\)', '', str(x)) for x in df.iloc[:, 0]]
                        cnts = [val_c(v) for v in df.iloc[:, 1]]
                        money = [val_c(v) for v in df.iloc[:, 2]]
                        b_bars = ax.bar(yrs, cnts, color='#4472C4', label='立项数(项)', width=0.5)
                        for b in b_bars:
                            ax.text(b.get_x() + b.get_width()/2, b.get_height(), f'{int(b.get_height())}', ha='center', va='bottom', color='black')
                        ax2 = ax.twinx()
                        ax2.plot(yrs, money, color='#ED7D31', marker='o', linewidth=2, label='到账经费(万元)')
                        for i, v in enumerate(money):
                            ax2.text(yrs[i], v, f'{v:.2f}', ha='center', va='bottom', color='black')
                        fig.legend(loc='lower center', bbox_to_anchor=(0.5, 0.02), ncol=2)
                        plt.subplots_adjust(top=0.88, bottom=0.15)

                    # 4. 横向年份趋势图 (图10, 11)
                    elif num_val in [10, 11]:
                        h_cols = [re.sub(r'[^\d]', '', str(h)) for h in df.columns]
                        yr_cols = [h for h in h_cols if len(h) == 4]
                        if len(yr_cols) >= 3:
                            x_list, y_list = yr_cols, [val_c(df.iloc[0, i]) for i in range(1, len(df.columns))]
                        else:
                            x_list, y_list = [re.sub(r'（.*?）|\(.*?\)', '', str(x)) for x in df.iloc[:, 0]], [val_c(v) for v in df.iloc[:, 1]]
                        bars = ax.bar(x_list, y_list, color='#4472C4', width=0.5)
                        for b in bars:
                            ax.text(b.get_x()+b.get_width()/2, b.get_height(), f'{int(b.get_height())}', ha='center', va='bottom', color='black')

                    # 5. 普通柱状图
                    else:
                        x_list = [str(x).replace('万元', '').strip() for x in df.iloc[:, 0]]
                        y_list = [val_c(v) for v in df.iloc[:, 1]]
                        bars = ax.bar(x_list, y_list, color='#4472C4', width=0.5)
                        for i, v in enumerate(y_list):
                            ax.text(i, v, f'{int(v)}', ha='center', va='bottom', color='black')

                    # 设置轴标签样式
                    lbls = {1:('发表年份','论文数量'), 3:('年份','立项数量'), 6:('经费区间（万元）','项目数量'), 7:('立项年份','项目数量'), 8:('经费区间','项目数量')}
                    if num_val in lbls:
                        ax.set_xlabel(lbls[num_val][0], fontweight='bold', color='black')
                        ax.set_ylabel(lbls[num_val][1], fontweight='bold', color='black')

                    buf = io.BytesIO(); plt.savefig(buf, format='png', dpi=150, bbox_inches='tight'); plt.close(); buf.seek(0)
                    doc.add_picture(buf, width=Inches(5.6))
                    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
                else:
                    # --- 表格逻辑 ---
                    table = doc.add_table(rows=len(raw_data), cols=len(header))
                    table.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    table.allow_autofit = False
                    
                    for i, r_cells in enumerate(raw_data):
                        table.rows[i].height = Cm(0.71)
                        row_txt = "".join(r_cells)
                        
                        # 【核心修改】：表 10 样式，B级/C级分类行背景修正为蓝色 (#D9E1F2)
                        is_rank = ("B级" in row_txt or "C级" in row_txt) and num_val == 10
                        
                        if is_rank:
                            m_cell = table.cell(i, 0).merge(table.cell(i, len(header)-1))
                            for p in m_cell.paragraphs: p.clear()
                            p = m_cell.paragraphs[0]
                            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            m_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                            set_font(p.add_run(row_txt.replace('*', '').strip()), 11, True)
                            set_cell_shading(m_cell, "D9E1F2") 
                        else:
                            for j, val in enumerate(r_cells):
                                if j >= len(header): break
                                cell = table.cell(i, j)
                                # 宽度适配
                                h_name = str(header[j])
                                if "序号" in h_name: cell.width = Cm(1.0)
                                elif "姓名" in h_name: cell.width = Cm(2.0)
                                elif any(x in h_name for x in ["单位", "学院"]): cell.width = Cm(4.8)
                                elif re.match(r'^20\d{2}$', val) or "年份" in h_name: cell.width = Cm(1.8)
                                else: cell.width = Cm(2.5)

                                cell.text = val
                                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                                p = cell.paragraphs[0]
                                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                                
                                if i == 0:
                                    set_cell_shading(cell, "4472C4") 
                                    set_font(p.runs[0] if p.runs else p.add_run(val), 11, True)
                                elif i % 2 == 0:
                                    set_cell_shading(cell, "D9E1F2") 
                                    set_font(p.runs[0] if p.runs else p.add_run(val), 11, False)
                                else:
                                    set_font(p.runs[0] if p.runs else p.add_run(val), 11, False)
                    set_table_border(table)
            except Exception as e:
                print(f"Error flushing table {current_title}: {e}")
        
        table_rows, is_chart_mode, current_title = [], False, ""

    # 循环遍历行内容
    for line in lines:
        clean_line = line.strip()
        if not clean_line: continue
        
        # 处理标题
        if clean_line.startswith('#'):
            flush_table()
            lvl = clean_line.count('#')
            if lvl <= 3:
                p = doc.add_heading('', level=lvl)
                p.paragraph_format.line_spacing = 1.5 # 1.5倍行距
                if "附录" in clean_line: p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                set_font(p.add_run(clean_line.replace('#', '').strip()), 14, True)
            else:
                p = doc.add_paragraph()
                p.paragraph_format.line_spacing = 1.5
                set_font(p.add_run(clean_line.replace('#', '').strip()), 12, True)
        
        # 处理图表题 (支持带星号或被center包裹清洗后的文本)
        elif re.match(r'(\*\*?)?(附)?[图表]\s?[\d\-\.]+[:：\s]', clean_line):
            flush_table()
            current_title = clean_line.replace('*', '').strip()
            is_chart_mode = "图" in current_title
            p = doc.add_paragraph()
            p.paragraph_format.line_spacing = 1.5
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_font(p.add_run(current_title), 11, True)
        
        # 处理表格内容
        elif clean_line.startswith('|'):
            table_rows.append(clean_line)
        
        # 处理普通段落
        else:
            if table_rows: flush_table()
            p = doc.add_paragraph()
            p.paragraph_format.line_spacing = 1.5 # 1.5倍行距
            p.paragraph_format.first_line_indent = Pt(24) 
            set_font(p.add_run(clean_line.replace('**', '')), 12)
            
    flush_table()

def add_toc(doc):
    """添加目录"""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.line_spacing = 1.5
    set_font(p.add_run("目  录"), size=16, bold=True)
    
    p_toc = doc.add_paragraph()
    p_toc.paragraph_format.line_spacing = 1.5
    run_toc = p_toc.add_run()
    f1 = OxmlElement('w:fldChar'); f1.set(qn('w:fldCharType'), 'begin')
    inst = OxmlElement('w:instrText'); inst.set(qn('xml:space'), 'preserve'); inst.text = 'TOC \\o "1-3" \\h \\z \\u'
    f2 = OxmlElement('w:fldChar'); f2.set(qn('w:fldCharType'), 'separate')
    hint = OxmlElement('w:t'); hint.text = "（请在此处右键单击，选择“更新域”以生成目录内容）"
    f3 = OxmlElement('w:fldChar'); f3.set(qn('w:fldCharType'), 'end')
    run_toc._r.extend([f1, inst, f2, hint, f3])
    doc.add_page_break()

@app.post("/generate_report_word")
async def generate_report_word(input_data: ReportInput, request: Request):
    try:
        doc = Document()
        doc.settings.element.append(OxmlElement('w:updateFields')).set(qn('w:val'), 'true')
        
        # 封面页
        for _ in range(4): 
            doc.add_paragraph().paragraph_format.line_spacing = 1.5
        
        p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.line_spacing = 1.5
        set_font(p.add_run("山东师范大学人文社会科学科研成果发展态势分析报告\n（2020-2024）"), size=22, bold=True)
        
        doc.add_page_break()
        add_toc(doc)
        add_page_number(doc)
        
        # 章节处理
        for i in range(1, 9): 
            text_content = getattr(input_data, f"ch{i}_text", "")
            if text_content and text_content.strip():
                process_smart(doc, text_content)
        
        # 附录处理
        if input_data.appendix and input_data.appendix.strip():
            doc.add_page_break()
            process_smart(doc, input_data.appendix)
            
        filename = f"report_{uuid.uuid4().hex[:8]}.docx"
        file_path = os.path.join("static", filename)
        doc.save(file_path)
        
        return {"file": f"{str(request.base_url).rstrip('/')}/static/{filename}", "status": "success"}
    except Exception as e:
        # 如果出错，返回错误信息以便调试
        return {"file": "", "status": "error", "message": str(e)}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8080)
