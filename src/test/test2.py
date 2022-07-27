from docx import Document
from docx.shared import Pt  #设置像素、缩进等
from docx.shared import RGBColor #设置字体颜色
from docx.oxml.ns import qn


doc = Document(r"E:\Projects\Project\PyPDF\CustomWatermark\inputs\那天晚上.docx")

for paragraph in doc.paragraphs:
    paragraph.paragraph_format.line_spacing = 1.5
    for run in paragraph.runs:
        run.font.bold = True
        run.font.italic = False
        run.font.underline = False
        run.font.strike = False
        run.font.shadow = False
        run.font.size = Pt(10.5)
        run.font.color.rgb = RGBColor(0,0,0)
        run.font.name = "等线"

        # 设置像黑体这样的中文字体，必须添加下面 2 行代码
        r = run._element.rPr.rFonts
        r.set(qn("w:eastAsia"),"等线")

doc.save(r"E:\Projects\Project\PyPDF\CustomWatermark\\那天晚上.docx")
