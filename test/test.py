from docx import Document
from docx_attach import replace_xlsx
from docx_attach import replace_docx


def test(): 
    doc = Document('template.docx')
    replace_docx(doc, 'word_attachment', 'word_attachment.docx')
    replace_xlsx(doc, 'excel_attachment', 'excel_attachment.xlsx')
    doc.save('new.docx')


if __name__ == '__main__':
    test()
