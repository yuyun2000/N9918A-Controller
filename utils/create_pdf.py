from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.platypus import Table, TableStyle, Paragraph
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
import os
import re

# === 固定公司信息 ===
COMPANY_NAME = "深圳市明栈信息科技有限公司"
ENG_COMPANY_NAME = "M5Stack Technology Co., Ltd"

# === 中文字体注册 ===
font_path = './simfang.ttf'
bold_font_path = './simhei.ttf'
song_font_path = './simsun.ttc'  # 宋体

try:
    if os.path.exists(font_path):
        pdfmetrics.registerFont(TTFont('simfang', font_path))
    else:
        pdfmetrics.registerFont(TTFont('simfang', 'C:/Windows/Fonts/simfang.ttf'))
except:
    print("警告：未找到仿宋字体文件")

try:
    if os.path.exists(bold_font_path):
        pdfmetrics.registerFont(TTFont('simhei', bold_font_path))
    else:
        pdfmetrics.registerFont(TTFont('simhei', 'C:/Windows/Fonts/simhei.ttf'))
except:
    print("警告：未找到黑体字体文件")

try:
    if os.path.exists(song_font_path):
        pdfmetrics.registerFont(TTFont('simsun', song_font_path))
    else:
        pdfmetrics.registerFont(TTFont('simsun', 'C:/Windows/Fonts/simsun.ttc'))
except:
    print("警告：未找到宋体字体文件")

def generate_test_report(
    filename="test_report.pdf",
    logo_path="../assets/m5logo2022.png",
    project_info=None,
    test_graph_path=None,
    spectrum_data=None,
    summary_text=None
):
    """
    生成测试报告PDF
    """
    # 默认项目信息
    if project_info is None:
        project_info = {}
    default_project_info = {
        'customer': 'N/A',
        'eut': 'N/A',
        'model': 'N/A',
        'mode': '工作模式',
        'engineer': 'Eden Chen',
        'remark': 'A2'
    }
    # 合并项目信息
    for key, value in default_project_info.items():
        if key not in project_info or not project_info[key]:
            project_info[key] = value

    # 创建PDF
    c = canvas.Canvas(filename, pagesize=A4)
    width, height = A4

    # 样式定义
    styleH = ParagraphStyle(
        'Heading',
        fontSize=16,
        leading=18,
        alignment=1,
        spaceAfter=6,
        spaceBefore=12,
        fontName='simhei'
    )
    styleTableHd = ParagraphStyle(
        'TableHd',
        fontSize=8,
        leading=10,
        alignment=1,
        fontName='simhei',
        textColor=colors.white
    )
    styleSectionTitle = ParagraphStyle(
        'SectionTitle',
        fontSize=11,
        leading=14,
        spaceAfter=10,
        spaceBefore=15,
        fontName='simhei'
    )

    # AI总结样式定义
    styleSummaryTitle = ParagraphStyle(
        'SummaryTitle',
        fontName='simhei',
        fontSize=14,
        leading=18,
        spaceAfter=20,
        spaceBefore=10,
        alignment=1  # 居中
    )

    # 第一页
    current_y = _draw_first_page(c, width, height, logo_path, project_info, test_graph_path, spectrum_data,
                                styleH, styleTableHd, styleSectionTitle)

    # 第二页 - 总结页
    c.showPage()
    _draw_summary_page(c, width, height, summary_text, styleSummaryTitle)

    c.save()
    print(f"PDF已生成: {filename}")

def _draw_first_page(c, width, height, logo_path, project_info, test_graph_path, spectrum_data,
                     styleH, styleTableHd, styleSectionTitle):
    """绘制第一页内容，返回当前Y坐标"""
    # Logo（缩小尺寸）
    logo_width = 50
    logo_height = 50
    if os.path.exists(logo_path):
        try:
            from PIL import Image as PILImage
            pil_img = PILImage.open(logo_path)
            aspect_ratio = pil_img.width / pil_img.height
            actual_height = logo_width / aspect_ratio
            c.drawImage(logo_path, 45, height - 70, width=logo_width, height=actual_height, mask='auto')
            logo_height = actual_height
        except Exception as e:
            print("加载图片失败:", e)
            c.drawImage(logo_path, 45, height - 70, width=logo_width, height=logo_height, mask='auto')

    # 水平分割线
    c.line(40, height - 90, width - 40, height - 90)

    # 标题
    p = Paragraph("Test Report", styleH)
    p.wrapOn(c, 400, 50)
    p.drawOn(c, (width - 400) / 2, height - 130)
    current_y = height - 140

    # 项目信息表标题
    section_title = Paragraph("Project Information", styleSectionTitle)
    section_title.wrapOn(c, 400, 20)
    section_title.drawOn(c, 50, current_y)
    current_y -= 10

    # 项目信息表（去除 SN/Voltage/Env/Test Info 等字段）
    data = [
        ['Customer:', project_info['customer'], 'EUT:', project_info['eut']],
        ['Model:', project_info['model'], 'Mode:', project_info['mode']],
        ['Engineer:', project_info['engineer'], 'Remark:', project_info['remark']]
    ]
    table = Table(data, colWidths=[70, 110, 70, 145])
    table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONT', (0, 0), (-1, -1), 'simfang'),
        ('FONTSIZE', (0, 0), (-1, -1), 9),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('FONTNAME', (0, 0), (0, -1), 'simhei'),
        ('FONTNAME', (2, 0), (2, -1), 'simhei'),
    ]))
    table.wrapOn(c, width - 100, 100)
    table.drawOn(c, 50, current_y - 50)
    current_y -= 70

    # 测试图表标题
    section_title3 = Paragraph("Test Graph", styleSectionTitle)
    section_title3.wrapOn(c, 400, 20)
    section_title3.drawOn(c, 50, current_y)
    current_y -= 20

    if test_graph_path and os.path.exists(test_graph_path):
        c.drawImage(test_graph_path, 55, current_y - 200, width=width - 110, height=200, preserveAspectRatio=True)
    else:
        c.setStrokeColor(colors.lightgrey)
        c.rect(55, current_y - 200, width - 110, 200)
    current_y -= 220

    # Suspected List 标题
    section_title4 = Paragraph("Suspected List", styleSectionTitle)
    section_title4.wrapOn(c, 400, 20)
    section_title4.drawOn(c, 50, current_y)
    current_y -= 15

    # 解析并绘制频谱数据表
    if spectrum_data:
        if isinstance(spectrum_data, list):
            table_data = _parse_spectrum_data_list(spectrum_data)
        else:
            lines = str(spectrum_data).split('\n')
            table_data = _parse_spectrum_data_list(lines)

        if table_data and len(table_data) > 1:
            # 自动计算列宽
            col_count = len(table_data[0])
            col_widths = [(width - 100) // col_count] * col_count  # 平均分布
            row_height = 18  # 更紧凑行距
            table_height = len(table_data) * row_height

            # 判断是否需要分页
            if current_y - table_height < 50:
                available_height = current_y - 80
                max_rows = max(1, int(available_height / row_height))
                if max_rows >= len(table_data):
                    _draw_table_on_page(c, table_data, col_widths, 50, current_y - table_height, row_height)
                else:
                    partial_data = table_data[:max_rows]
                    _draw_table_on_page(c, partial_data, col_widths, 50, current_y - (len(partial_data) * row_height), row_height)
                    remaining_data = table_data[max_rows:]
                    if remaining_data:
                        c.showPage()
                        _draw_logo_only_header(c, width, height, logo_path)
                        current_y = height - 120
                        _draw_table_on_page(c, remaining_data, col_widths, 50, current_y - (len(remaining_data) * row_height), row_height)
            else:
                _draw_table_on_page(c, table_data, col_widths, 50, current_y - table_height, row_height)
        else:
            print("未解析到有效表格数据")
    else:
        print("未提供频谱数据")

    return current_y

def _draw_logo_only_header(c, width, height, logo_path):
    """只绘制Logo的页头"""
    logo_width = 50
    logo_height = 50
    if os.path.exists(logo_path):
        try:
            from PIL import Image as PILImage
            pil_img = PILImage.open(logo_path)
            aspect_ratio = pil_img.width / pil_img.height
            actual_height = logo_width / aspect_ratio
            c.drawImage(logo_path, 45, height - 70, width=logo_width, height=actual_height, mask='auto')
        except Exception as e:
            print("加载图片失败:", e)
            c.drawImage(logo_path, 45, height - 70, width=logo_width, height=logo_height, mask='auto')
    c.line(40, height - 90, width - 40, height - 90)

def _draw_table_on_page(c, table_data, col_widths, x, y, row_height):
    """在指定位置绘制表格"""
    num_rows = len(table_data)
    # 单独设置第一行高度更大一些（例如30），其它维持默认row_height
    row_heights = [30] + [row_height] * (num_rows - 1)  # 👈 表头用大一点的高度
    
    table = Table(table_data, colWidths=col_widths, rowHeights=row_heights)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#797d80')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'simhei'),
        ('FONTSIZE', (0, 0), (-1, 0), 8),
        ('FONTNAME', (0, 1), (-1, -1), 'simfang'),
        ('FONTSIZE', (0, 1), (-1, -1), 8),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('WORDWRAP', (0, 0), (-1, -1)),
    ]))
    total_table_height = sum(row_heights)
    table.wrapOn(c, sum(col_widths), total_table_height)
    table.drawOn(c, x, y)

def _parse_spectrum_data_list(spectrum_input):
    """解析频谱数据，支持字符串输入，仅保留前15条有效数据"""
    if not spectrum_input:
        return None

    # 如果是字符串，则按行拆分
    if isinstance(spectrum_input, str):
        lines = spectrum_input.strip().split('\n')
    else:
        lines = [line.strip() for line in spectrum_input if line.strip()]

    if len(lines) < 3:
        return None

    # 定义表头
    table_header = [
        'NO.',
        'Freq\n[MHz]',
        'Amplitude\n[dBuV]',  # 修复：用 u 替换 μ
        'FCC Limit\n[dBuV]',  # 修复：用 u 替换 μ
        'FCC Margin\n[dB]',
        'Status'
    ]
    table_data = [table_header]
    count = 0  # 记录有效数据行数

    for line in lines:
        line = line.strip()
        # 忽略标题行、分隔线等无效内容
        if 'No' in line and 'Freq' in line:
            continue
        if line.startswith('-') or '----' in line:
            continue
        if not line:
            continue

        parts = line.split()
        if len(parts) >= 6:
            table_data.append(parts[:6])
            count += 1
            if count >= 15:  # 只保留前15个点
                break

    return table_data if len(table_data) > 1 else None

def _clean_text_for_pdf(text):
    """清理文本中可能导致显示问题的字符 - 修复版"""
    # 替换各种可能导致问题的字符
    replacements = {
        # 直接删除或替换问题字符
        '•': '●',      # bullet point -> 实心圆点
        '–': '-',      # en dash  
        '—': '-',      # em dash
        '"': '"',      # left double quotation
        '"': '"',      # right double quotation
        ''': "'",      # left single quotation
        ''': "'",      # right single quotation
        '…': '...',    # ellipsis
        '×': 'x',      # multiplication sign
        '°': '度',     # degree symbol
        'μ': 'u',      # micro symbol
        '★': '*',      # 星号替换
        '☆': '*',      # 空心星号替换
    }
    
    for old, new in replacements.items():
        text = text.replace(old, new)
    
    return text

def _parse_markdown_content(text):
    """解析Markdown内容，返回结构化数据 - 修复版"""
    if not text:
        return []
    
    # 清理文本
    text = _clean_text_for_pdf(text)
    lines = text.strip().split('\n')
    
    content_blocks = []
    
    for line in lines:
        line = line.strip()
        
        if not line:
            content_blocks.append({'type': 'space', 'content': ''})
            continue
    
        # H3标题 (###)
        if line.startswith('#### '):
            content_blocks.append({
                'type': 'h4',
                'content': line[5:].strip()
            })
        
        # H3标题 (###)
        if line.startswith('### '):
            content_blocks.append({
                'type': 'h3',
                'content': line[4:].strip()
            })
        # H2标题 (##)
        elif line.startswith('## '):
            content_blocks.append({
                'type': 'h2', 
                'content': line[3:].strip()
            })
        # H1标题 (#)
        elif line.startswith('# '):
            content_blocks.append({
                'type': 'h1',
                'content': line[2:].strip()
            })
        # 数字列表 (1. 2. 3.)
        elif re.match(r'^\d+\.\s+', line):
            match = re.match(r'^(\d+)\.\s+(.*)', line)
            if match:
                num, content = match.groups()
                content_blocks.append({
                    'type': 'ordered_list',
                    'number': num,
                    'content': content
                })
        # 无序列表 (* - +) - 修复：统一使用 ● 符号
        elif line.startswith(('* ', '- ', '+ ')):
            content_blocks.append({
                'type': 'unordered_list',
                'content': line[2:].strip()  # 去掉前面的符号
            })
        # 普通段落（可能包含粗体）
        else:
            content_blocks.append({
                'type': 'paragraph',
                'content': line
            })
    
    return content_blocks

def _process_bold_text(text):
    """处理粗体标记 - 修复版"""
    # 将 **text** 转换为 <font name="simhei">text</font>
    # 这样可以确保粗体文本正确显示
    processed_text = re.sub(r'\*\*(.*?)\*\*', r'<font name="simhei">\1</font>', text)
    return processed_text

def _draw_summary_page(c, width, height, summary_text, styleTitle):
    """绘制AI总结页面 - 修复版"""
    
    # 绘制页面标题
    title = Paragraph("AI测试分析报告", styleTitle)
    title.wrapOn(c, width - 100, 50)
    title.drawOn(c, 50, height - 80)
    
    if not summary_text:
        return
    
    current_y = height - 120
    margin_left = 50
    margin_right = 50
    content_width = width - margin_left - margin_right
    
    # 解析Markdown内容
    content_blocks = _parse_markdown_content(summary_text)
    
    # 定义各种样式
    styles = {
        'h1': ParagraphStyle(
            'H1Style',
            fontName='simhei',
            fontSize=13,
            leading=16,
            spaceAfter=12,
            spaceBefore=15,
            textColor=colors.HexColor('#2c3e50')
        ),
        'h2': ParagraphStyle(
            'H2Style', 
            fontName='simhei',
            fontSize=12,
            leading=15,
            spaceAfter=10,
            spaceBefore=12,
            textColor=colors.HexColor('#34495e')
        ),
        'h3': ParagraphStyle(
            'H3Style',
            fontName='simhei', 
            fontSize=11,
            leading=14,
            spaceAfter=8,
            spaceBefore=10,
            textColor=colors.HexColor('#5d6d7e')
        ),
        'paragraph': ParagraphStyle(
            'ParagraphStyle',
            fontName='simfang',
            fontSize=10,
            leading=13,
            spaceAfter=6,
            leftIndent=0,
            textColor=colors.black
        ),
        'ordered_list': ParagraphStyle(
            'OrderedListStyle',
            fontName='simfang',
            fontSize=10,
            leading=13,
            spaceAfter=4,
            leftIndent=15,
            bulletIndent=0,
            textColor=colors.black
        ),
        'unordered_list': ParagraphStyle(
            'UnorderedListStyle',
            fontName='simfang',
            fontSize=10,
            leading=13,
            spaceAfter=4,
            leftIndent=15,
            bulletIndent=0,
            textColor=colors.black
        )
    }
    
    # 渲染每个内容块
    for block in content_blocks:
        block_type = block['type']
        content = block.get('content', '')
        
        # 空行处理
        if block_type == 'space':
            current_y -= 6
            continue
        
        # 检查是否需要换页
        estimated_height = 30  # 预估高度
        if current_y - estimated_height < 60:
            c.showPage()
            current_y = height - 60
        
        # 根据类型渲染内容
        if block_type in ['h1', 'h2', 'h3']:
            style = styles[block_type]
            p = Paragraph(content, style)
            w, h = p.wrap(content_width, 100)
            p.drawOn(c, margin_left, current_y - h)
            current_y -= h + style.spaceAfter
            
        elif block_type == 'ordered_list':
            number = block.get('number', '1')
            style = styles['ordered_list']
            # 处理粗体
            processed_content = _process_bold_text(content)
            formatted_content = f"{number}. {processed_content}"
            p = Paragraph(formatted_content, style)
            w, h = p.wrap(content_width - 15, 200)
            p.drawOn(c, margin_left, current_y - h)
            current_y -= h + style.spaceAfter
            
        elif block_type == 'unordered_list':
            style = styles['unordered_list']
            # 处理粗体
            processed_content = _process_bold_text(content)
            # 使用实心圆点 ●，确保能正常显示
            formatted_content = f"● {processed_content}"
            p = Paragraph(formatted_content, style)
            w, h = p.wrap(content_width - 15, 200)
            p.drawOn(c, margin_left, current_y - h)
            current_y -= h + style.spaceAfter
            
        else:  # 普通段落
            style = styles['paragraph']
            # 处理粗体
            processed_content = _process_bold_text(content)
            p = Paragraph(processed_content, style)
            w, h = p.wrap(content_width, 200)
            p.drawOn(c, margin_left, current_y - h)
            current_y -= h + style.spaceAfter

# 使用示例
if __name__ == "__main__":
    # 示例频谱数据（用户提供的格式）
    spectrum_data = '''
QUASI_PEAK Mode Results:
====================================================================================================
No   Freq [MHz]   Amplitude [dBμV]   FCC Limit [dBμV]   FCC Margin [dB]    Status         
----------------------------------------------------------------------------------------------------
1    175.015      42.82              40.0               2.82               FCC Fail       
2    274.925      47.79              46.0               1.79               FCC Fail       
3    46.975       39.91              40.0               -0.09              Pass           
4    224.970      44.75              46.0               -1.25              Pass           
5    499.965      38.77              46.0               -7.23              Pass           
6    76.075       31.28              40.0               -8.72              Pass           
7    240.005      36.50              46.0               -9.50              Pass           
8    159.980      27.64              40.0               -12.36             Pass           
9    72.680       27.63              40.0               -12.37             Pass           
10   52.795       26.01              40.0               -13.99             Pass           
11   450.010      31.46              46.0               -14.54             Pass           
12   350.100      31.41              46.0               -14.59             Pass           
13   170.650      24.65              40.0               -15.35             Pass           
14   64.435       24.60              40.0               -15.40             Pass           
15   69.285       24.26              40.0               -15.74             Pass           
'''
    
    # 示例总结文本
    summary_text = """### 异常频点及简要数据信息列表  
1. **46.975 MHz**：Amplitude=39.91 dBμV，限值=40.0 dBμV，Margin=-0.09 dB，Status=Pass（临界，接近限值）
2. **175.015 MHz**：Amplitude=42.82 dBμV，限值=40.0 dBμV，Margin=2.82 dB，Status=Fail（超出限值2.82 dB）
3. **274.925 MHz**：Amplitude=47.79 dBμV，限值=46.0 dBμV，Margin=1.79 dB，Status=Fail（超出限值1.79 dB，Margin≤2 dB临界）

### 异常点间的内在规律性
1. **25 MHz基准时钟谐波关联**：175.015 MHz（25 MHz×7）和274.925 MHz（25 MHz×11）精确对应25 MHz基准时钟的7次、11次谐波（理论值分别为175 MHz、275 MHz），频率偏差<0.1 MHz，高度符合时钟谐波序列特征。
2. **低频临界频点关联性**：46.975 MHz接近25 MHz×1.88（约47 MHz），可能为25 MHz时钟的2次谐波（50 MHz）的偏移，或与25 MHz时钟源相关的低频干扰（如电源纹波、晶振寄生频率）。

## 测试建议

* 检查25MHz时钟信号的屏蔽效果
* 优化电源设计以减少纹波干扰
* 考虑添加滤波器来抑制谐波辐射

## 总结

该产品在EMC测试中表现出明显的时钟谐波问题，需要针对性的设计改进。
"""
    
    # 项目信息
    project_info = {
        'customer': 'M5Stack',
        'eut': '产品A',
        'model': 'Model-X',
        'mode': '正常工作模式',
        'engineer': '张工程师',
        'remark': '首次测试'
    }
    
    # 生成报告
    generate_test_report(
        filename="test_report.pdf",
        logo_path="../assets/m5logo2022.png",
        project_info=project_info,
        test_graph_path="../assets/m5logo2022.png",
        spectrum_data=spectrum_data,
        summary_text=summary_text
    )