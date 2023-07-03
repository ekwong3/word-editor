import sys
import os
import string
from docx import Document


def read_doc(doc):
    print('reading', doc)
    document = Document(doc)
    for section in document.sections:
        for p in section.first_page_header.paragraphs:
            print(p.text)
        for p in section.header.paragraphs:
            print(p.text)
        for p in section.even_page_header.paragraphs:
            print(p.text)

    for paragraph in document.paragraphs:
        print(paragraph.text)

    for section in document.sections:
        for p in section.first_page_footer.paragraphs:
            print(p.text)
        for p in section.footer.paragraphs:
            print(p.text)
        for p in section.even_page_footer.paragraphs:
            print(p.text)


def is_punc(char):
    return char in string.punctuation + string.whitespace


def matches_word(old_text, index, word_length):
    if (index == 0):
        return is_punc(old_text[word_length])
    elif (index + word_length == len(old_text)):
        return is_punc(old_text[index - 1])
    else:
        return is_punc(old_text[index - 1]) and is_punc(old_text[index + word_length])


def get_new_word(old_word, replace_word):
    if (old_word.islower()):
        return replace_word.lower()
    elif (old_word.isupper()):
        return replace_word.upper()
    elif (old_word[0].isupper() and old_word[1:].islower()):
        return replace_word[0].upper() + replace_word[1:].lower()
    else:
        return replace_word


def get_new_text(old_text, find_word, replace_word, keep_case):
    word_length = len(find_word)
    text_copy = old_text.lower()
    find_copy = find_word.lower()
    res = ""
    ind = text_copy.find(find_copy)
    while ind >= 0:
        if matches_word(text_copy, ind, word_length):
            res += old_text[:ind]
            if (keep_case):
                old_word = old_text[ind:ind + word_length]
                res += get_new_word(old_word, replace_word)
            else:
                res += replace_word
        else:
            res += old_text[:ind + word_length]
        text_copy = text_copy[ind + word_length:]
        old_text = old_text[ind + word_length:]
        ind = text_copy.find(find_copy)
    res += old_text
    return res


def find_and_replace(doc, find, replace, keep_case):
    document = Document(doc)
    sections = document.sections
    for i in range(len(document.paragraphs)):
        paragraph_text = document.paragraphs[i].text
        new_text = get_new_text(paragraph_text, find, replace, keep_case)
        document.paragraphs[i].text = new_text

    for section in sections:
        if section.different_first_page_header_footer:
            diff_header = section.first_page_header
            for i in range(len(diff_header.paragraphs)):
                header_text = diff_header.paragraphs[i].text
                new_text = get_new_text(header_text, find, replace, keep_case)
                diff_header.paragraphs[i].text = new_text

            diff_footer = section.first_page_footer
            for i in range(len(diff_footer.paragraphs)):
                footer_text = diff_footer.paragraphs[i].text
                new_text = get_new_text(footer_text, find, replace, keep_case)
                diff_footer.paragraphs[i].text = new_text

        if document.settings.odd_and_even_pages_header_footer:
            diff_header = section.even_page_header
            for i in range(len(diff_header.paragraphs)):
                header_text = diff_header.paragraphs[i].text
                new_text = get_new_text(header_text, find, replace, keep_case)
                diff_header.paragraphs[i].text = new_text

            diff_footer = section.even_page_footer
            for i in range(len(diff_footer.paragraphs)):
                footer_text = diff_footer.paragraphs[i].text
                new_text = get_new_text(footer_text, find, replace, keep_case)
                diff_footer.paragraphs[i].text = new_text

        header = section.header
        for i in range(len(header.paragraphs)):
            header_text = header.paragraphs[i].text
            new_text = get_new_text(header_text, find, replace, keep_case)
            header.paragraphs[i].text = new_text

        footer = section.footer
        for i in range(len(footer.paragraphs)):
            footer_text = footer.paragraphs[i].text
            new_text = get_new_text(footer_text, find, replace, keep_case)
            footer.paragraphs[i].text = new_text

    document.save(doc)


def replace_many(doc, finds, replaces, keep_case):
    finds = finds.split('\n')
    replaces = replaces.split('\n')
    shorter = min(len(finds), len(replaces))
    for i in range(shorter):
        find_and_replace(doc, finds[i], replaces[i], keep_case)


def replace_folder(folderpath, finds, replaces, keep_case, process_sub):
    for path in os.listdir(folderpath):
        if path.endswith('.docx'):
            filepath = f'{folderpath}{path}'
            replace_many(filepath, finds, replaces, keep_case)
        elif process_sub and os.path.isdir(f'{folderpath}{path}/'):
            replace_folder(f'{folderpath}{path}/', finds,
                           replaces, keep_case, process_sub)


folder = sys.argv[1]
find = sys.argv[2]
replace = sys.argv[3]
keep_case = sys.argv[4] == "true"
process_sub = sys.argv[5] == "true"

replace_folder(folder, find, replace, keep_case, process_sub)
