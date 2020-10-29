import os
import openpyxl

from io import StringIO
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage

import xInvoices as xI
from common import program_objs as objs


def filename_from_path(path):
    return list(map(lambda x: x.split("\\"), path.split("/")))[-1][-1]


def listdir_by_ext(path, extension):
    files = os.listdir(path)

    files = [check_ext(file, extension) for file in files]

    files = tuple([f"{xI.PDFS_PATH}/{file}" for file in files])

    return files


def check_ext(file, extension):
    if file.endswith(extension):
        return file


def read_excel(xlsx):
    return openpyxl.load_workbook(xlsx)


# https://stackoverflow.com/questions/26494211/extracting-text-from-a-pdf-file-using-pdfminer-in-python
def read_pdf(path):
    rsrcmgr = PDFResourceManager()
    retstr = StringIO()
    laparams = LAParams()
    device = TextConverter(rsrcmgr, retstr, laparams=laparams)
    fp = open(path, 'rb')
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    password = ""
    maxpages = 0
    caching = True
    pagenos = set()

    for page in PDFPage.get_pages(fp, pagenos, maxpages=maxpages, password=password, caching=caching,
                                  check_extractable=True):
        interpreter.process_page(page)

    text = retstr.getvalue()

    fp.close()
    device.close()
    retstr.close()

    return text


if __name__ == '__main__':
    pass
