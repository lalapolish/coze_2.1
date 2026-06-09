import re
import io
import pandas as pd
import matplotlib.pyplot as plt
from docx import Document
from docx.shared import Pt, Inches
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH

# --- 绘图配置与清理函数（同前，保持不变） ---
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
    headers = [re.sub(r'\$', '', c).strip() for c in lines[0].split('|') if c.strip()]
    data = []
    for line in lines[2:]:
        row = [clean_data(c) for c in line.split('|') if c.strip()]
        if len(row) == len(headers): data.append(row)
    return pd.DataFrame(data, columns=headers)

# --- 绘图逻辑（根据你的要求区分样式） ---
def generate_chart(df, title, fig_no):
    plt.figure(figsize=(10, 6))
    img_stream = io.BytesIO()
    
    try:
        if fig_no in [2, 12]: # 分组柱状图（参考上传图样式）
            df.set_index(df.columns[0]).plot(kind='bar', color=GROUPED_COLORS, ax=plt.gca(), width=0.8)
            plt.legend(bbox_to_anchor=(0.5, -0.2), loc='upper center', ncol=5, frameon=False)
        
        elif fig_no in [4, 5]: # 饼图
            plt.pie(df.iloc[:, -1], labels=df.iloc[:, 0], autopct='%1.1f%%', colors=plt.cm.Pastel1.colors)
        
        elif fig_no == 9: # 柱状图+折线图（双轴）
            fig, ax1 = plt.subplots(figsize=(10, 6))
            x = df.iloc[:, 0].astype(str)
            ax1.bar(x, df.iloc[:, 1], color=BLUE_COLOR, label='项目数量')
            ax2 = ax1.twinx()
            ax2.plot(x, df.iloc[:, 2], color='#ED7D31', marker='o', linewidth=2, label='到账经费')
            fig.legend(loc="upper right", bbox_to_anchor=(0.9, 0.9))
        
        else: # 普通蓝色柱状图
            bars = plt.bar(df.iloc[:, 0].astype(str), df.iloc[:, 1], color=BLUE_COLOR)
            for bar in bars:
                plt.text(bar.get_x() + bar.get_width()/2, bar.get_height(), f'{bar.get_height()}', ha='center', va='bottom')
        
        plt.title(title, fontsize=14, pad=20)
        plt.tight_layout()
        plt.savefig(img_stream, format='png', dpi=300)
    finally:
        plt.close()
    return img_stream

# --- Word 样式设置 ---
def set_style(obj):
    """设置宋体小四"""
    if hasattr(obj, 'runs'):
        for run in obj.runs:
            run.font.size = Pt(12)
            run.font.name = '宋体'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

# --- 核心处理函数 ---
def process_content(doc, full_text):
    # 分解文字、标题和表格
    parts = re.split(r'(\*\*图 \d+：.*?\*\*|\*\*表 \d+：.*?\*\*|## .*?|### .*?)', full_text)
    current_fig_title = None

    for part in parts:
        part = part.strip()
        if not part: continue

        if part.startswith('##'): # 章标题
            h = doc.add_heading(part.replace('#','').strip(), level=2)
            set_style(h)
        elif part.startswith('###'): # 节标题
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
            
            if current_fig_title: # 如果是图，绘图并插入
                fig_no_match = re.search(r'图 (\d+)', current_fig_title)
                fig_no = int(fig_no_match.group(1)) if fig_no_match else 0
                img = generate_chart(df, current_fig_title, fig_no)
                doc.add_picture(img, width=Inches(5.8))
                doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
                current_fig_title = None
            else: # 如果是表，直接画表格
                table = doc.add_table(rows=1, cols=len(df.columns))
                table.style = 'Table Grid'
                for i, col in enumerate(df.columns):
                    table.rows[0].cells[i].text = str(col)
                for _, row in df.iterrows():
                    row_cells = table.add_row().cells
                    for i, val in enumerate(row):
                        row_cells[i].text = str(val)
        else: # 纯文字
            if "import matplotlib" in part or "![" in part: continue
            p = doc.add_paragraph(part.replace('$', ''))
            set_style(p)

# --- 插件入口函数 (工作流核心) ---
def main(args):
    """
    args 是一个字典，key 对应你在工作流节点设置的输入变量名
    """
    doc = Document()
    
    # 按照 2-7 章的顺序合并内容
    # 假设你在输入里定义了 ch2, ch3, ch4, ch5, ch6, ch7
    for i in range(2, 8):
        key = f"ch{i}_text"
        chapter_content = args.get(key, "") # 获取每一章的内容
        if chapter_content:
            process_content(doc, chapter_content)
    
    # 保存文档
    output_path = "Scientific_Report.docx"
    doc.save(output_path)
    
    # 在工作流中，通常需要返回文件的 URL 或者路径
    return {
        "word_file": output_path,
        "status": "success"
    }
