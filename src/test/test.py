# from PyPDF2 import PdfFileWriter, PdfFileReader


# 给pdf批量加水印
# input_pdf = 'y8.PDF',输入文件
# output = 'y8_q1.pdf',输出pdf
# watermark = 's.pdf'水印文件


# def create_watermark(input_pdf, output, watermark):
#     watermark_obj = PdfFileReader(watermark)
#     watermark_page = watermark_obj.getPage(0)
#     pdf_reader = PdfFileReader(input_pdf)
#     pdf_writer = PdfFileWriter()
#     # 给所有页面添加水印
#     for page_num in range(pdf_reader.getNumPages()):
#         print("page:", page_num)
#         page = pdf_reader.getPage(page_num)
#         if page_num % 2 == 0:
#             page.mergePage(watermark_page)
#         pdf_writer.addPage(page)
#     ###这一行是加密，如果只想加密，上面的添加水印都可以删除。
#     pdf_writer.encrypt(user_pwd="", owner_pwd="xx1234")  # 设置pdf密码
#     # pdf_writer.
#     # pdf_writer.encrypt()
#     with open(output, 'wb') as out:
#         pdf_writer.write(out)

from hashlib import md5

from PyPDF4 import PdfFileReader, PdfFileWriter
from PyPDF4.generic import NameObject, DictionaryObject, ArrayObject, \
    NumberObject, ByteStringObject
from PyPDF4.pdf import _alg33, _alg34, _alg35
from PyPDF4.utils import b_


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


unmeta = PdfFileReader('E:\Projects\Project\PyPDF\CustomWatermark\\result\那天晚上.pdf')

writer = PdfFileWriter()
writer.appendPagesFromReader(unmeta)
encrypt(writer, '', '123')

with open('E:\Projects\Project\PyPDF\CustomWatermark\\那天晚上.pdf', 'wb') as fp:
    writer.write(fp)

# create_watermark("E:\Projects\Project\PyPDF\CustomWatermark\\result\那天晚上.pdf","E:\Projects\Project\PyPDF\CustomWatermark\\那天晚上.pdf","E:\Projects\Project\PyPDF\CustomWatermark\\result\watermark.pdf")