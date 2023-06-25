from docx import Document

document = Document('Test.docx')
sections = document.sections
header = sections[0].header
header.paragraphs[0].text = 'changing the header to this'
print(header.paragraphs[0].text)
document.save('Test.docx')
