import os
import fitz
from shutil import rmtree
import img2pdf
from PyQt5.QtWidgets import QMessageBox
from pandas import read_excel
from pathlib import Path
from reportlab.lib.units import cm
from reportlab.pdfgen import canvas
from win32com.client import DispatchEx
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from PyPDF2 import PdfFileWriter, PdfFileReader
from hashlib import md5
from PyPDF4 import PdfFileReader, PdfFileWriter
from PyPDF4.generic import NameObject, DictionaryObject, ArrayObject, NumberObject, ByteStringObject
from PyPDF4.pdf import _alg33, _alg34, _alg35
from PyPDF4.utils import b_
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import LAParams, LTText
from ProgressBar import pyqtbar


class AddWaterMask(object):

    def __init__(self, input_file_path, output_file_path):
        pdfmetrics.registerFont(TTFont('kaiti', 'C:/Windows/Fonts/simkai.ttf'))  # 楷体
        self.input_file_path = input_file_path
        self.output_file_path = output_file_path
        word_exist = len([x for x in os.listdir(input_file_path) if x.endswith('.docx') or x.endswith('.doc')])
        excel_exist = len([x for x in os.listdir(input_file_path) if x.endswith('.xlsx') or x.endswith('.xls')])
        if word_exist == 0 or excel_exist == 0:
            QMessageBox.critical(None, "错误", "所选输入文件所在的文件夹未包含word文档或excel文档")
            self.flag = 0
        else:
            self.flag = 1
            self.TEMP_DIR = os.path.join(os.getcwd(), 'temp')
            self.OUTPUT_DIR = output_file_path
            self.INPUT_DIR = input_file_path
            if not os.path.exists(self.TEMP_DIR):
                os.mkdir(self.TEMP_DIR)
            if not os.path.exists(self.OUTPUT_DIR):
                os.mkdir(self.OUTPUT_DIR)
            if not os.path.exists(self.INPUT_DIR):
                os.mkdir(self.INPUT_DIR)
            self.word_abs_path = os.path.join(input_file_path ,[x for x in os.listdir(input_file_path) if
                                                                x.endswith('.docx') or x.endswith('.doc')][0])
            self.excel_abs_path = os.path.join(input_file_path , [x for x in os.listdir(input_file_path) if
                                                            x.endswith('.xlsx') or x.endswith('.xls')][0])
            self.temp_word_abs_path = os.path.join(self.TEMP_DIR,os.path.basename(self.word_abs_path))
            self.temp_pdf_abs_path = os.path.join(self.TEMP_DIR,
                                             os.path.basename(self.word_abs_path).replace('.docx', '.pdf').replace('.doc',
                                                                                                                   '.pdf'))
            self.persons = read_excel(self.excel_abs_path).to_dict('records')

    def convert_word2pdf(self):
        app = DispatchEx('Word.Application')
        app.Visible = 0
        app.DisplayAlerts = 0
        doc = app.Documents.Open(self.word_abs_path)
        # raise ValueError('A very specific bad thing happened.')
        doc.SaveAs(self.temp_pdf_abs_path, FileFormat=17)
        doc.Close()
        app.Quit()

    def create_watermark(self,content,target_path):
        fp = open(target_path, 'rb')
        parser = PDFParser(fp)
        doc: PDFDocument = PDFDocument(parser)
        rsrcmgr = PDFResourceManager()
        laparams = LAParams()
        device = PDFPageAggregator(rsrcmgr, laparams=laparams)
        interpreter = PDFPageInterpreter(rsrcmgr, device)

        space = 0.825
        value_est1 = 763.665
        value_est2 = 740.265
        space_raw = value_est1 - value_est2
        space_factor = space_raw / space
        page_num = 0
        for page in PDFPage.create_pages(doc):
            c = canvas.Canvas(os.path.join(self.TEMP_DIR, 'watermark-%i.pdf') % page_num, pagesize=(21 * cm, 29.7 * cm))
            c.setFillColorRGB(190 / 255, 190 / 255, 190 / 255, alpha=0.4)  # 淡
            page_num = page_num + 1
            c.translate(0 * cm, 27 * cm)  # 移动坐标原点(坐标系左下为(0,0)))
            c.setFont('kaiti', 10.5)
            interpreter.process_page(page)
            layout = device.get_result()
            pre = 790
            cnt = 0
            for textbox in layout:
                if isinstance(textbox, LTText):
                    for line in textbox:
                        cur = line.bbox[3]
                        if (pre - cur > 22):
                            if cnt % 3 == 0:
                                c.setFillColorRGB(190 / 255, 190 / 255, 190 / 255, alpha=0.4)  # 淡
                                c.setFont('kaiti', 10.5)
                                if line.bbox[0] < 340:
                                    c.drawString(3 * cm, ((line.bbox[3] - value_est1) / space_factor) * cm, content)
                                    c.drawString(8.5 * cm, ((line.bbox[3] - value_est1) / space_factor) * cm, content)
                                c.drawString(14 * cm, ((line.bbox[3] - value_est1) / space_factor) * cm, content)
                            cnt = cnt + 1
                        if line.width < 300 and (pre - cur > 22) and line.bbox[0] < 120 and line.width > 20:
                            c.setFillColorRGB(190 / 255, 190 / 255, 190 / 255, alpha=1)  # 淡
                            c.setFont('kaiti', 6)
                            c.drawString((3 + (line.width + line.bbox[0] - 90) / 26.5) * cm,
                                         ((line.bbox[3] - value_est1) / space_factor - space / 2) * cm, content)
                        if pre > cur:
                            pre = cur
            c.save()
        fp.close()

    def add_watermark2pdf(self,input_pdf, output_pdf, TEMP_DIR):
        pdf = PdfFileReader(input_pdf, strict=False)
        pdf_writer = PdfFileWriter()
        page_num = 0
        for page in range(pdf.getNumPages()):
            watermark_pdf = os.path.join(TEMP_DIR, 'watermark-%i.pdf') % page_num
            page_num = page_num + 1
            watermark = PdfFileReader(watermark_pdf)
            watermark_page = watermark.getPage(0)
            pdf_page = pdf.getPage(page)
            pdf_page.mergePage(watermark_page)
            pdf_page.compressContentStreams()
            pdf_writer.addPage(pdf_page)
        pdfOutputFile = open(output_pdf, 'wb')
        pdf_writer.write(pdfOutputFile)
        pdfOutputFile.close()

    def encrypt(self, writer_obj: PdfFileWriter, user_pwd, owner_pwd=None, use_128bit=True):
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

    def run(self):
        if self.flag == 1:
            ##
            bar = pyqtbar()
            total_number = len(self.persons)
            task_id = 1
            self.convert_word2pdf()
            bar.set_value(task_id, total_number, 5)
            for person in self.persons:
                bar.set_value(task_id, total_number, 10)
                wtmk_content = '仅供%s-%s参考' % (person['fund_company'], person['reseacher'])
                input_pdf = self.temp_pdf_abs_path
                output_pdf = os.path.join(self.TEMP_DIR, os.path.splitext(os.path.basename(self.word_abs_path))[0] + '_' + person[
                    'fund_company'] + '_' + person['reseacher'] + "_tmp" + '.pdf')
                self.create_watermark(wtmk_content,input_pdf)
                bar.set_value(task_id, total_number, 30)
                self.add_watermark2pdf(input_pdf, output_pdf, self.TEMP_DIR)
                bar.set_value(task_id, total_number, 60)
                # To get better resolution
                zoom_x = 4.0  # horizontal zoom
                zoom_y = 4.0  # vertical zoom
                mat = fitz.Matrix(zoom_x, zoom_y)  # zoom factor 2 in each dimension
                doc = fitz.open(output_pdf)  # open document
                for page in doc:
                    pix = page.get_pixmap(matrix=mat)  # render page to an image
                    pix.save(os.path.join(self.TEMP_DIR,"page-%i.png") % page.number)  # store image as a PNG
                doc.close()
                bar.set_value(task_id, total_number, 80)
                output_pdf = os.path.join(self.TEMP_DIR, os.path.splitext(os.path.basename(self.word_abs_path))[0] + '_' + person[
                    'fund_company'] + '_' + person['reseacher'] + '.pdf')
                with open(output_pdf, "wb") as f:
                    f.write(img2pdf.convert([str(path) for path in Path(self.TEMP_DIR).glob('*.png')]))
                    f.close()
                bar.set_value(task_id, total_number, 93)
                unmeta = PdfFileReader(output_pdf,strict=False)
                writer = PdfFileWriter()
                writer.appendPagesFromReader(unmeta)
                bar.set_value(task_id, total_number, 96)
                self.encrypt(writer, '', '123')
                final_pdf = os.path.join(self.OUTPUT_DIR, os.path.splitext(os.path.basename(self.word_abs_path))[0] + '_' + person[
                    'fund_company'] + '_' + person['reseacher'] + '.pdf')
                with open(final_pdf, 'wb') as fp:
                    writer.write(fp)
                bar.set_value(task_id, total_number, 99)
                task_id = task_id + 1
            rmtree(self.TEMP_DIR)
