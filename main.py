import re
import io
import pandas as pd
import matplotlib.pyplot as plt
from docx import Document
from docx.shared import Pt, Inches
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH

# 设置绘图字体（防止中文乱码，建议使用微软雅黑或黑体）
plt.rcParams['font.sans-serif'] = ['SimHei'] 
plt.rcParams['axes.unicode_minus'] = False

# --- 核心配置 ---
BLUE_COLOR = 'royalblue'
GROUPED_COLORS = ['#4472C4', '#ED7D31', '#FFC000', '#70AD47', '#25B6C7'] # 对应 B, C, D, E, F 等级颜色

def clean_data(text):
    """清除大模型输出中的 $$ 符号、逗号、百分号，并转为数字"""
    if pd.isna(text): return 0
    clean_val = str(text).replace('$', '').replace(',', '').replace('%', '').strip()
    try:
        return float(clean_val) if '.' in clean_val else int(clean_val)
    except:
        return clean_val

def md_table_to_df(md_text):
    """将 Markdown 表格转换为 Pandas DataFrame"""
    lines = [line.strip() for line in md_text.strip().split('\n') if '|' in line]
    if len(lines) < 2: return None
    # 提取表头和数据（跳过分隔行）
    headers = [re.sub(r'\$', '', c).strip() for c in lines[0].split('|') if c.strip()]
    data = []
    for line in lines[2:]:
        row = [clean_data(c) for c in line.split('|') if c.strip()]
        if len(row) == len(headers):
            data.append(row)
    return pd.DataFrame(data, columns=headers)

# --- 绘图函数集 ---

def draw_bar_chart(df, title, filename):
    """通用蓝色柱状图"""
    plt.figure(figsize=(10, 6))
    x_col = df.columns[0]
    y_col = df.columns[1]
    bars = plt.bar(df[x_col], df[y_col], color=BLUE_COLOR)
    for bar in bars:
        height = bar.get_height()
        plt.text(bar.get_x() + bar.get_width()/2, height, f'{height}', ha='center', va='bottom')
    plt.title(title)
    plt.tight_layout()
    img_stream = io.BytesIO()
    plt.savefig(img_stream, format='png', dpi=300)
    plt.close()
    return img_stream

def draw_grouped_bar(df, title, filename):
    """分组柱状图（图2, 图12专用颜色）"""
    plt.figure(figsize=(12, 6))
    df.set_index(df.columns[0]).plot(kind='bar', color=GROUPED_COLORS, ax=plt.gca())
    plt.title(title)
    plt.legend(bbox_to_anchor=(0.5, -0.15), loc='upper center', ncol=5)
    plt.tight_layout()
    img_stream = io.BytesIO()
    plt.savefig(img_stream, format='png', dpi=300)
    plt.close()
    return img_stream

def draw_pie_chart(df, title, filename):
    """饼图（图4, 图5专用）"""
    plt.figure(figsize=(8, 8))
    # 假设第一列是标签，最后一列是数值
    labels = df.iloc[:, 0]
    values = df.iloc[:, -1]
    plt.pie(values, labels=labels, autopct='%1.1f%%', colors=plt.cm.Paired.colors)
    plt.title(title)
    img_stream = io.BytesIO()
    plt.savefig(img_stream, format='png', dpi=300)
    plt.close()
    return img_stream

def draw_combo_chart(df, title, filename):
    """柱状图+折线图（图9专用）"""
    fig, ax1 = plt.subplots(figsize=(10, 6))
    x = df.iloc[:, 0].astype(str)
    y_bar = df.iloc[:, 1] # 项目数量
    y_line = df.iloc[:, 2] # 到账经费

    ax1.bar(x, y_bar, color=BLUE_COLOR, label='项目数量')
    ax1.set_ylabel('项目数量（项）')
    
    ax2 = ax1.twinx()
    ax2.plot(x, y_line, color='darkorange', marker='o', linewidth=2, label='到账经费')
    ax2.set_ylabel('到账经费（万元）')
    
    plt.title(title)
    fig.legend(loc="upper right", bbox_to_anchor=(1,1), bbox_transform=ax1.transAxes)
    img_stream = io.BytesIO()
    plt.savefig(img_stream, format='png', dpi=300)
    plt.close()
    return img_stream

# --- Word 处理逻辑 ---

def set_font_style(run):
    """设置宋体小四"""
    run.font.size = Pt(12) # 小四是12pt
    run.font.name = '宋体'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

def main_process(full_text):
    doc = Document()
    
    # 1. 预处理文本：按行切分并提取
    # 识别标题和表格
    parts = re.split(r'(\*\*图 \d+：.*?\*\*|\*\*表 \d+：.*?\*\*|## .*?|### .*?)', full_text)
    
    current_fig_title = None
    
    for part in parts:
        part = part.strip()
        if not part: continue

        # 处理标题（章、节）
        if part.startswith('#'):
            level = part.count('#')
            title_text = part.replace('#', '').strip()
            h = doc.add_heading(title_text, level=level)
            for run in h.runs:
                set_font_style(run)
        
        # 处理 图 X 标题
        elif "**图" in part:
            current_fig_title = part.strip("*")
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(current_fig_title)
            run.bold = True
            set_font_style(run)

        # 处理 表 X 标题
        elif "**表" in part:
            current_fig_title = None # 此时不是图了
            p = doc.add_paragraph()
            run = p.add_run(part.strip("*"))
            run.bold = True
            set_font_style(run)

        # 处理 表格数据内容
        elif part.startswith('|'):
            df = md_table_to_df(part)
            if df is None: continue

            # 如果上方刚出现了“图 X”，则转换成图
            if current_fig_title:
                fig_no = re.search(r'图 (\d+)', current_fig_title)
                img_stream = None
                if fig_no:
                    n = int(fig_no.group(1))
                    if n in [2, 12]:
                        img_stream = draw_grouped_bar(df, current_fig_title, n)
                    elif n in [4, 5]:
                        img_stream = draw_pie_chart(df, current_fig_title, n)
                    elif n == 9:
                        img_stream = draw_combo_chart(df, current_fig_title, n)
                    else:
                        img_stream = draw_bar_chart(df, current_fig_title, n)
                
                if img_stream:
                    doc.add_picture(img_stream, width=Inches(5.5))
                    last_p = doc.paragraphs[-1]
                    last_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                current_fig_title = None # 图画完了
            
            # 否则（或者它是表 X），在 Word 里画表格
            else:
                table = doc.add_table(rows=1, cols=len(df.columns))
                table.style = 'Table Grid'
                # 表头
                hdr_cells = table.rows[0].cells
                for i, col in enumerate(df.columns):
                    hdr_cells[i].text = str(col)
                # 数据
                for _, row in df.iterrows():
                    row_cells = table.add_row().cells
                    for i, val in enumerate(row):
                        row_cells[i].text = str(val)
        
        # 处理 普通分析文字
        else:
            # 过滤掉 Python 代码块和图片链接预览
            if "import matplotlib" in part or "plt.show()" in part or "![" in part:
                continue
            p = doc.add_paragraph(part.replace('$', '')) # 去掉文字里的 $$
            for run in p.runs:
                set_font_style(run)

    doc.save("学术评估报告_自动生成.docx")
    print("报告已生成！")

# --- 测试运行 ---
# 将你刚才贴出的那段大模型输出（包含文字和表格）赋值给 input_text
input_text = """
## 第2章 发文规模分析
...此处放你大模型输出的全文内容...
"""

main_process(input_text)
