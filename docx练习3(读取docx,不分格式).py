from docx import Document

dir_docx = '卫辉市省级内业核查意见.docx'

dd = Document(dir_docx)

for p in dd.paragraphs:

    print (p.text)
