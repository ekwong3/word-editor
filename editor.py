from docx import Document

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

def read_doc(doc):
    document = Document(doc)
    for p in document.sections[0].header.paragraphs:
        print(p.text)

    for paragraph in document.paragraphs:
        print(paragraph.text)
    
    for p in document.sections[0].footer.paragraphs:
        print(p.text)

# read_doc('Test.docx')