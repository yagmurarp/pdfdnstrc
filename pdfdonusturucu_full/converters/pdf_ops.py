from pdf2docx import Converter

def pdf_to_docx(src_pdf, dst_docx):
    cv = Converter(src_pdf)
    cv.convert(dst_docx, start=0, end=None)
    cv.close()
