import os
from win32com.client import DispatchEx
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from PyPDF2 import PdfFileWriter, PdfFileReader
import pandas as pd
from docx import Document
from docx.shared import Pt  #设置像素、缩进等
from docx.shared import RGBColor #设置字体颜色
from docx.oxml.ns import qn
from hashlib import md5
from PyPDF4 import PdfFileReader, PdfFileWriter
from PyPDF4.generic import NameObject, DictionaryObject, ArrayObject, \
    NumberObject, ByteStringObject
from PyPDF4.pdf import _alg33, _alg34, _alg35
from PyPDF4.utils import b_
import shutil
import glob, fitz
from pathlib import Path
import img2pdf
from typing import Union, Tuple
from reportlab.lib import units
from typing import List
from pikepdf import Pdf, Page, Rectangle

pdfmetrics.registerFont(TTFont('kaiti', 'C:/Windows/Fonts/simkai.ttf'))  # 楷体
pdfmetrics.registerFont(TTFont('msyh', r'./msyh.ttc'))

TEMP_DIR = os.path.join(os.getcwd(), 'temp')
MINUTE_DIR = os.path.join(os.getcwd(), 'result')
INPUT_DIR = os.path.join(os.getcwd(), 'inputs')
if not os.path.exists(TEMP_DIR):
    os.mkdir(TEMP_DIR)
if not os.path.exists(MINUTE_DIR):
    os.mkdir(MINUTE_DIR)
if not os.path.exists(INPUT_DIR):
    os.mkdir(INPUT_DIR)

word_abs_path = os.path.join(os.getcwd(), 'inputs', [x for x in os.listdir(os.path.join(os.getcwd(), 'inputs')) if
                                                     x.endswith('.docx') or x.endswith('.doc')][0])
excel_abs_path = os.path.join(os.getcwd(), 'inputs', [x for x in os.listdir(os.path.join(os.getcwd(), 'inputs')) if
                                                      x.endswith('.xlsx') or x.endswith('.xls')][0])


def convert_word2pdf(word_abs_path):
    app = DispatchEx('Word.Application')
    app.Visible = 0  # 这个至少在调试阶段建议打开，否则如果等待时间长的话，它至少给你耐心。。。
    app.DisplayAlerts = 0
    doc = app.Documents.Open(word_abs_path)

    all_content = doc.Range(doc.Content.Start, doc.Content.End)
    # all_content.HighlightColorIndex = 16  # 全局背景色淡黄色

    temp_pdf_abs_path = os.path.join(TEMP_DIR,
                                     os.path.basename(word_abs_path).replace('.docx', '.pdf').replace('.doc', '.pdf'))

    doc.SaveAs(temp_pdf_abs_path, FileFormat=17)
    doc.Close()
    app.Quit()
    return temp_pdf_abs_path  # 返回临时pdf的路径


####### 1.生成水印pdf的函数 ########
def create_watermark(content):
    # 默认大小为21cm*29.7cm
    c = canvas.Canvas(os.path.join(TEMP_DIR, 'watermark.pdf'), pagesize=(21 * cm, 29.7 * cm))
    c.translate(0 * cm, 27 * cm)  # 移动坐标原点(坐标系左下为(0,0)))
    c.setFont('kaiti', 10.5)

    c.setFillColorRGB(190 / 255, 190 / 255, 190 / 255, alpha=1)  # 淡
    space = -0.825
    # c.drawString(3 * cm, 0 * cm, content)
    # c.drawString(3 * cm, space * cm, content)
    # c.drawString(3 * cm, space * 2 * cm, content)
    # c.drawString(3 * cm, space * 3 * cm, content)
    # c.drawString(3 * cm, space * 4 * cm, content)
    # c.drawString(3 * cm, space * 5 * cm, content)
    # c.drawString(14 * cm, 0 * cm, content)
    # c.drawString(14 * cm, space * cm, content)
    # c.drawString(14 * cm, space * 2 * cm, content)
    # c.drawString(14 * cm, space * 3 * cm, content)
    # c.drawString(14 * cm, space * 4 * cm, content)
    # c.drawString(14 * cm, space * 5 * cm, content)
    for i in [3,8.5,14]:
        for j in range(0,32,2):
            c.drawString(i * cm, space * j * cm, content)
    c.save()  # 关闭并保存pdf文件

######## 2.为pdf文件加水印的函数 ########
def add_watermark2pdf(input_pdf, output_pdf, watermark_pdf):
    watermark = PdfFileReader(watermark_pdf)
    watermark_page = watermark.getPage(0)
    pdf = PdfFileReader(input_pdf, strict=False)
    pdf_writer = PdfFileWriter()
    for page in range(pdf.getNumPages()):
        pdf_page = pdf.getPage(page)
        pdf_page.mergePage(watermark_page)
        # watermark_page.mergePage(pdf_page)
        pdf_page.compressContentStreams()
        pdf_writer.addPage(pdf_page)
    pdfOutputFile = open(output_pdf, 'wb')
    pdf_writer.write(pdfOutputFile)
    pdfOutputFile.close()

def formatWord(input_word_path,output_word_path):
    doc = Document(input_word_path)
    for paragraph in doc.paragraphs:
        paragraph.paragraph_format.line_spacing = 1.5
        for run in paragraph.runs:
            run.font.bold = True
            run.font.italic = False
            run.font.underline = False
            run.font.strike = False
            run.font.shadow = False
            run.font.size = Pt(10.5)
            run.font.color.rgb = RGBColor(0, 0, 0)
            run.font.name = "等线"

            # 设置像黑体这样的中文字体，必须添加下面 2 行代码
            r = run._element.rPr.rFonts
            r.set(qn("w:eastAsia"), "等线")
    doc.save(output_word_path)

def encrypt(writer_obj: PdfFileWriter, user_pwd, owner_pwd=None, use_128bit=True):
    """
    Encrypt this PDF file with the PDF Standard encryption handler.

    :param str user_pwd: The "user password", which allows for opening
        and reading the PDF file with the restrictions provided.
    :param str owner_pwd: The "owner password", which allows for
        opening the PDF files without any restrictions.  By default,
        the owner password is the same as the user password.
    :param bool use_128bit: flag as to whether to use 128bit
        encryption.  When false, 40bit encryption will be used.  By default,
        this flag is on.
    """
    import time, random
    if owner_pwd == None:
        owner_pwd = user_pwd
    if use_128bit:
        V = 2
        rev = 3
        keylen = int(128 / 8)
    else:
        V = 1
        rev = 2
        keylen = int(40 / 8)
    # permit copy and printing only:
    P = -3904
    O = ByteStringObject(_alg33(owner_pwd, user_pwd, rev, keylen))
    ID_1 = ByteStringObject(md5(b_(repr(time.time()))).digest())
    ID_2 = ByteStringObject(md5(b_(repr(random.random()))).digest())
    writer_obj._ID = ArrayObject((ID_1, ID_2))
    if rev == 2:
        U, key = _alg34(user_pwd, O, P, ID_1)
    else:
        assert rev == 3
        U, key = _alg35(user_pwd, rev, keylen, O, P, ID_1, False)
    encrypt = DictionaryObject()
    encrypt[NameObject("/Filter")] = NameObject("/Standard")
    encrypt[NameObject("/V")] = NumberObject(V)
    if V == 2:
        encrypt[NameObject("/Length")] = NumberObject(keylen * 8)
    encrypt[NameObject("/R")] = NumberObject(rev)
    encrypt[NameObject("/O")] = ByteStringObject(O)
    encrypt[NameObject("/U")] = ByteStringObject(U)
    encrypt[NameObject("/P")] = NumberObject(P)
    writer_obj._encrypt = writer_obj._addObject(encrypt)
    writer_obj._encrypt_key = key


def create_watermark2(content: str,
                     filename: str,
                     width: Union[int, float],
                     height: Union[int, float],
                     font: str,
                     fontsize: int,
                     angle: Union[int, float] = 45,
                     text_stroke_color_rgb: Tuple[int, int, int] = (0, 0, 0),
                     text_fill_color_rgb: Tuple[int, int, int] = (0, 0, 0),
                     text_fill_alpha: Union[int, float] = 1) -> None:
    '''
    用于生成包含content文字内容的水印pdf文件
    content: 水印文本内容
    filename: 导出的水印文件名
    width: 画布宽度，单位：mm
    height: 画布高度，单位：mm
    font: 对应注册的字体代号
    fontsize: 字号大小
    angle: 旋转角度
    text_stroke_color_rgb: 文字轮廓rgb色
    text_fill_color_rgb: 文字填充rgb色
    text_fill_alpha: 文字透明度
    '''

    # 创建pdf文件，指定文件名及尺寸，这里以像素单位为例
    c = canvas.Canvas(filename, pagesize=(width * units.mm, height * units.mm))

    # 进行轻微的画布平移保证文字的完整
    c.translate(0.1 * width * units.mm, 0.1 * height * units.mm)

    # 设置旋转角度
    c.rotate(angle)

    # 设置字体及字号大小
    c.setFont(font, fontsize)

    # 设置文字轮廓色彩
    c.setStrokeColorRGB(*text_stroke_color_rgb)

    # 设置文字填充色
    c.setFillColorRGB(*text_fill_color_rgb)

    # 设置文字填充色透明度
    c.setFillAlpha(text_fill_alpha)

    # 绘制文字内容
    c.drawString(4, 4, content)

    # 保存水印pdf文件
    c.save()


def add_watermark2(target_pdf_path: str,
                  watermark_pdf_path: str,
                  output_pdf_path: str,
                  nrow: int,
                  ncol: int,
                  skip_pages: List[int] = []) -> None:
    '''
    向目标pdf文件中添加平铺水印
    target_pdf_path: 目标pdf文件的路径+文件名
    watermark_pdf_path: 水印pdf文件的路径+文件名
    nrow: 水印平铺的行数
    ncol：水印平铺的列数
    skip_pages: 需要跳过不添加水印的页面序号（从0开始）
    '''

    # 读入需要添加水印的pdf文件
    target_pdf = Pdf.open(target_pdf_path)

    # 读入水印pdf文件并提取水印页
    watermark_pdf = Pdf.open(watermark_pdf_path)
    watermark_page = watermark_pdf.pages[0]

    # 遍历目标pdf文件中的所有页（排除skip_pages指定的若干页）
    for idx, target_page in enumerate(target_pdf.pages):

        if idx not in skip_pages:
            for x in range(ncol):
                for y in range(nrow):
                    # 向目标页指定范围添加水印
                    target_page.add_overlay(watermark_page, Rectangle(target_page.trimbox[2] * x / ncol,
                                                                      target_page.trimbox[3] * y / nrow,
                                                                      target_page.trimbox[2] * (x + 1) / ncol,
                                                                      target_page.trimbox[3] * (y + 1) / nrow))

    # 将添加完水印后的结果保存为新的pdf
    target_pdf.save(output_pdf_path)


def main():
    print('程序正在运行…………by Superon')
    ### temp路径
    temp_word_abs_path = os.path.join(TEMP_DIR,os.path.basename(word_abs_path))

    ### 格式化word
    formatWord(word_abs_path,temp_word_abs_path)

    ### word转pdf
    temp_pdf_abs_path = convert_word2pdf(temp_word_abs_path)

    ### 获取人员名单
    persons = pd.read_excel(excel_abs_path).to_dict('records')

    for person in persons:
        ### 创建水印PDF
        wtmk_content = '仅供%s-%s参考' % (person['fund_company'], person['reseacher'])
        print(wtmk_content)
        # create_watermark(wtmk_content)

        ### 合并纪要pdf和水印pdf
        watermark_pdf = os.path.join(TEMP_DIR, 'watermark.pdf')
        create_watermark2(content=wtmk_content,
                          filename=watermark_pdf,
                          width=150,
                          height=150,
                          font='msyh',
                          fontsize=30,
                          text_fill_alpha=0.5)

        input_pdf = temp_pdf_abs_path
        output_pdf = os.path.join(TEMP_DIR, os.path.splitext(os.path.basename(word_abs_path))[0] + '_' + person[
            'fund_company'] + '_' + person['reseacher'] + "_tmp" + '.pdf')
        # add_watermark2pdf(input_pdf, output_pdf, watermark_pdf)
        add_watermark2(target_pdf_path=input_pdf,
                      watermark_pdf_path=watermark_pdf,
                      output_pdf_path=output_pdf,
                      nrow=6,
                      ncol=4,
                      skip_pages=[0])


        # To get better resolution
        zoom_x = 4.0  # horizontal zoom
        zoom_y = 4.0  # vertical zoom
        mat = fitz.Matrix(zoom_x, zoom_y)  # zoom factor 2 in each dimension
        doc = fitz.open(output_pdf)  # open document
        for page in doc:  # iterate through the pages
            pix = page.get_pixmap(matrix=mat)  # render page to an image
            pix.save(os.path.join(TEMP_DIR,"page-%i.png") % page.number)  # store image as a PNG
        doc.close()

        output_pdf = os.path.join(TEMP_DIR, os.path.splitext(os.path.basename(word_abs_path))[0] + '_' + person[
            'fund_company'] + '_' + person['reseacher'] + '.pdf')
        with open(output_pdf, "wb") as f:
            f.write(img2pdf.convert([str(path) for path in Path(TEMP_DIR).glob('*.png')]))
            f.close()

        ### 加权限
        unmeta = PdfFileReader(output_pdf,strict=False)
        writer = PdfFileWriter()
        writer.appendPagesFromReader(unmeta)
        encrypt(writer, '', '123')
        final_pdf = os.path.join(MINUTE_DIR, os.path.splitext(os.path.basename(word_abs_path))[0] + '_' + person[
            'fund_company'] + '_' + person['reseacher'] + '.pdf')
        with open(final_pdf, 'wb') as fp:
            writer.write(fp)

    shutil.rmtree(TEMP_DIR)


if __name__ == '__main__':
    main()