import os
from win32com.client import DispatchEx
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics 
from reportlab.pdfbase.ttfonts import TTFont 
pdfmetrics.registerFont(TTFont('kaiti', 'C:/Windows/Fonts/simkai.ttf'))#楷体
from PyPDF2 import PdfFileWriter,PdfFileReader
import pandas as pd
import shutil

TEMP_DIR = os.path.join(os.getcwd(),'temp')
MINUTE_DIR = os.path.join(os.getcwd(),'minutes')
INPUT_DIR  = os.path.join(os.getcwd(),'inputs')
if not os.path.exists(TEMP_DIR):
    os.mkdir(TEMP_DIR)
if not os.path.exists(MINUTE_DIR):
    os.mkdir(MINUTE_DIR)
if not os.path.exists(INPUT_DIR):
    os.mkdir(INPUT_DIR)

word_abs_path = os.path.join(os.getcwd(),'inputs',[x for x in os.listdir(os.path.join(os.getcwd(),'inputs')) if x.endswith('.docx') or x.endswith('.doc')][0])
excel_abs_path = os.path.join(os.getcwd(),'inputs',[x for x in os.listdir(os.path.join(os.getcwd(),'inputs')) if x.endswith('.xlsx') or x.endswith('.xls')][0])

def convert_word2pdf(word_abs_path):
    app = DispatchEx('Word.Application')
    app.Visible = 0 #这个至少在调试阶段建议打开，否则如果等待时间长的话，它至少给你耐心。。。
    app.DisplayAlerts = 0
    doc = app.Documents.Open(word_abs_path)

    all_content = doc.Range(doc.Content.Start, doc.Content.End)
    all_content.HighlightColorIndex  = 16 #全局背景色淡黄色

    temp_pdf_abs_path = os.path.join(TEMP_DIR,os.path.basename(word_abs_path).replace('.docx','.pdf').replace('.doc','.pdf'))

    doc.SaveAs(temp_pdf_abs_path, FileFormat=17)
    doc.Close()
    app.Quit()
    return temp_pdf_abs_path #返回临时pdf的路径


####### 1.生成水印pdf的函数 ########
def create_watermark(content):
    #默认大小为21cm*29.7cm
    c = canvas.Canvas(os.path.join(TEMP_DIR,'watermark.pdf'), pagesize = (21*cm, 29.7*cm))   
    c.translate(10*cm, 15*cm) #移动坐标原点(坐标系左下为(0,0)))      
    c.setFont('kaiti',15)

    c.setFillColorRGB(190/255,190/255,190/255,alpha=0.4)#淡
    for i in range(-8,8):
        for j in range(-30,30):
            c.drawString((0.5+i*6)*cm, (0.5+j*0.6)*cm, content)                                                                                                                        
    c.save()#关闭并保存pdf文件

######## 2.为pdf文件加水印的函数 ########
def add_watermark2pdf(input_pdf,output_pdf,watermark_pdf):
    watermark = PdfFileReader(watermark_pdf)
    watermark_page = watermark.getPage(0)
    pdf = PdfFileReader(input_pdf,strict=False)
    pdf_writer = PdfFileWriter()
    for page in range(pdf.getNumPages()):
        pdf_page = pdf.getPage(page)
        pdf_page.mergePage(watermark_page)
        pdf_page.compressContentStreams() 
        pdf_writer.addPage(pdf_page)
    pdfOutputFile = open(output_pdf,'wb')   
    pdf_writer.write(pdfOutputFile)
    pdfOutputFile.close()

def main():
    print('程序正在运行…………by Superon')
    ### word转pdf
    temp_pdf_abs_path = convert_word2pdf(word_abs_path)

    ### 获取人员名单
    persons = pd.read_excel(excel_abs_path).to_dict('records')

    for person in persons:
        ### 创建水印PDF
        wtmk_content = '仅供%s-%s参考'%(person['fund_company'],person['reseacher'])
        print(wtmk_content)
        create_watermark(wtmk_content)
        ### 合并纪要pdf和水印pdf
        watermark_pdf = os.path.join(TEMP_DIR,'watermark.pdf')
        input_pdf = temp_pdf_abs_path
        output_pdf = os.path.join(TEMP_DIR,os.path.splitext(os.path.basename(word_abs_path))[0] + '_' + person['fund_company'] + '_' + person['reseacher'] + '.pdf')
        add_watermark2pdf(input_pdf,output_pdf,watermark_pdf)
        
        ### 权限设置
        final_pdf = os.path.join(MINUTE_DIR,os.path.splitext(os.path.basename(word_abs_path))[0] + '_' + person['fund_company'] + '_' + person['reseacher'] + '.pdf')
        # os.system("""pdftk "%s" output "%s" owner_pw 15263748"""%(output_pdf,final_pdf))
    # shutil.rmtree(TEMP_DIR)
        
if __name__=='__main__':   
    main()