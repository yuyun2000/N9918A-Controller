from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.platypus import Table, TableStyle, Paragraph
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
import os

# === 固定公司信息 ===
COMPANY_NAME = "深圳市明栈信息科技有限公司"
ENG_COMPANY_NAME = "M5Stack Technology Co., Ltd"

# === 中文字体注册 ===
font_path = './simfang.ttf'
bold_font_path = './simhei.ttf'

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

    styleSummaryTitle = ParagraphStyle(
        'SummaryTitle',
        fontName='simhei',
        fontSize=14,
        leading=18,
        spaceAfter=18,
        spaceBefore=18,
        alignment=1
    )

    styleSummaryContent = ParagraphStyle(
        'SummaryContent',
        fontName='simfang',
        fontSize=10,
        leading=14,
        leftIndent=20,
        spaceAfter=10
    )

    # 第一页
    current_y = _draw_first_page(c, width, height, logo_path, project_info, test_graph_path, spectrum_data,
                                styleH, styleTableHd, styleSectionTitle)

    # 第二页 - 总结页
    c.showPage()
    _draw_summary_page(c, width, height, summary_text, styleSummaryTitle, styleSummaryContent)

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

def _parse_spectrum_data_list(spectrum_lines):
    """解析频谱数据字符串列表"""
    if not spectrum_lines:
        return None

    lines = [line.strip() for line in spectrum_lines if line.strip()]
    if len(lines) < 3:
        return None

    # 创建带换行的表头
    table_header = [
        'NO.',
        'Freq\n[MHz]',
        'Amplitude\n[dBμV]',
        'FCC Limit\n[dBμV]',
        'FCC Margin\n[dB]',
        'Status'
    ]

    table_data = [table_header]

    for i, line in enumerate(lines):
        if 'No' in line and 'Freq' in line:
            continue
        if line.startswith('-') or '----' in line:
            continue
        if not line:
            continue

        parts = line.split()
        if len(parts) >= 6:
            table_data.append(parts[:6])

    return table_data if len(table_data) > 1 else None


def _draw_summary_page(c, width, height, summary_text, styleTitle, styleContent):
    """绘制总结页"""

    # 标题
    title = Paragraph("AI测试报告", styleTitle)
    title.wrapOn(c, width - 100, 50)
    title.drawOn(c, 50, height - 100)

    if summary_text:
        lines = summary_text.strip().split('\n')
        current_y = height - 130
        c.setFont('simfang', 10)

        for line in lines:
            if line.strip():
                if line.strip().startswith('*') or line.strip().startswith('-') or line.strip().startswith('  '):
                    c.drawString(70, current_y, line.strip())
                else:
                    c.setFont('simhei', 10)  # 加粗标题
                    c.drawString(50, current_y, line.strip())
                    c.setFont('simfang', 10)
                current_y -= 12

                if current_y < 100:
                    break

# 使用示例
if __name__ == "__main__":
    # 示例频谱数据（用户提供的格式）
    spectrum_data = [
        "No   Freq [MHz]   Amplitude [dBμV]   FCC Limit [dBμV]   FCC Margin [dB]    Status         ",
        "----------------------------------------------------------------------------------------------------",
        "1    175.015      43.35              40.0               3.35               FCC Fail       ",
        "2    274.925      48.06              46.0               2.06               FCC Fail       ",
        "3    47.945       39.34              40.0               -0.66              Pass           ",
        "4    224.970      44.93              46.0               -1.07              Pass           ",
        "5    240.005      40.48              46.0               -5.52              Pass     ",
        "6    499.965      38.91              46.0               -7.09              Pass     ",
        "7    499.965      38.91              46.0               -7.09              Pass     ",
        "8    499.965      38.91              46.0               -7.09              Pass     ",
        "9    499.965      38.91              46.0               -7.09              Pass     ",
        "10    499.965      38.91              46.0               -7.09              Pass     ",
        "11    499.965      38.91              46.0               -7.09              Pass     ",
        "12    499.965      38.91              46.0               -7.09              Pass     ",
        "13    499.965      38.91              46.0               -7.09              Pass     ",
        "14    499.965      38.91              46.0               -7.09              Pass     ",
        "15    499.965      38.91              46.0               -7.09              Pass     "
    ]
    
    # 示例总结文本
    summary_text = """一、概览
* 检测频段：30MHz-1GHz
* 测试采样时长：15秒
* 检测模式及采样点数：

二、问题点明细
以下为详细问题点分析...
测试结果显示，在175.015MHz和274.925MHz频率点存在超标现象，需要进一步分析整改。
其他频率点均符合FCC标准要求。"""
    
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