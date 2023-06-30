import sys
from docx import Document


def read_doc(doc):
    document = Document(doc)
    for p in document.sections[0].header.paragraphs:
        print(p.text)

    for paragraph in document.paragraphs:
        print(paragraph.text)

    for p in document.sections[0].footer.paragraphs:
        print(p.text)


def find_and_replace(doc, find, replace):
    document = Document(doc)
    sections = document.sections
    for i in range(len(document.paragraphs)):
        paragraph_text = document.paragraphs[i].text
        new_text = paragraph_text.replace(find, replace)
        document.paragraphs[i].text = new_text

    for section in sections:
        header = section.header
        for i in range(len(header.paragraphs)):
            header_text = header.paragraphs[i].text
            new_text = header_text.replace(find, replace)
            header.paragraphs[i].text = new_text

        footer = section.footer
        for i in range(len(footer.paragraphs)):
            footer_text = footer.paragraphs[i].text
            new_text = footer_text.replace(find, replace)
            footer.paragraphs[i].text = new_text

    document.save(doc)


path = sys.argv[1]
find = sys.argv[2]
replace = sys.argv[3]

read_doc(path)
find_and_replace(path, find, replace)
read_doc(path)
