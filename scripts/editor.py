import sys
import os
import string
from docx import Document

NON_PUNCTUATION = string.ascii_letters + string.digits


def read_doc(doc):
    print('reading', doc)
    document = Document(doc)
    print('reading headers')
    for section in document.sections:
        for p in section.first_page_header.paragraphs:
            print(p.text)
        for p in section.header.paragraphs:
            print(p.text)
        for p in section.even_page_header.paragraphs:
            print(p.text)

    print('reading paragraphs')
    for paragraph in document.paragraphs:
        print(paragraph.text)

    print('reading tables')
    for table in document.tables:
        for row in range(len(table.rows)):
            for col in range(len(table.columns)):
                print(table.cell(row, col).text)

    print('reading footers')
    for section in document.sections:
        for p in section.first_page_footer.paragraphs:
            print(p.text)
        for p in section.footer.paragraphs:
            print(p.text)
        for p in section.even_page_footer.paragraphs:
            print(p.text)


def is_punc(char):
    return char.lower() not in NON_PUNCTUATION


def matches_word(old_text, index, word_length):
    if (index + word_length == len(old_text) and index == 0):
        return True
    elif (index + word_length == len(old_text)):
        return is_punc(old_text[index - 1])
    elif (index == 0):
        return is_punc(old_text[word_length])
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


def get_new_text(old_text, find_word, replace_word, keep_case, match_word):
    word_length = len(find_word)
    text_copy = old_text.lower()
    find_copy = find_word.lower()
    res = ""
    ind = text_copy.find(find_copy)
    while ind >= 0:
        if (match_word and matches_word(text_copy, ind, word_length)) or (not match_word):
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


def find_and_replace_section(section, find, replace, keep_case, match_word):
    for i in range(len(section.paragraphs)):
        for run in section.paragraphs[i].runs:
            old_text = run.text
            new_text = get_new_text(
                old_text, find, replace, keep_case, match_word)
            run.text = new_text


def find_and_replace(doc, find, replace, keep_case, match_word):
    document = Document(doc)
    sections = document.sections
    find_and_replace_section(document, find, replace, keep_case, match_word)

    for section in sections:
        header = section.header
        find_and_replace_section(header, find, replace, keep_case, match_word)

        footer = section.footer
        find_and_replace_section(footer, find, replace, keep_case, match_word)

        if section.different_first_page_header_footer:
            diff_header = section.first_page_header
            find_and_replace_section(
                diff_header, find, replace, keep_case, match_word)

            diff_footer = section.first_page_footer
            find_and_replace_section(
                diff_footer, find, replace, keep_case, match_word)

        if document.settings.odd_and_even_pages_header_footer:
            diff_header = section.even_page_header
            find_and_replace_section(
                diff_header, find, replace, keep_case, match_word)

            diff_footer = section.even_page_footer
            find_and_replace_section(
                diff_footer, find, replace, keep_case, match_word)

    for table in document.tables:
        for row in range(len(table.rows)):
            for col in range(len(table.columns)):
                find_and_replace_section(table.cell(
                    row, col), find, replace, keep_case, match_word)

    document.save(doc)


def replace_many(doc, finds, replaces, keep_case, match_word):
    finds = finds.split('\n')
    replaces = replaces.split('\n')
    shorter = min(len(finds), len(replaces))
    for i in range(shorter):
        find_and_replace(doc, finds[i], replaces[i], keep_case, match_word)


def replace_folder(folderpath, finds, replaces, keep_case, match_word, process_sub):
    for path in os.listdir(folderpath):
        # omit non word files and temporary open ones
        if path.endswith('.docx') and path[:2] != '~$':
            filepath = f'{folderpath}{path}'
            replace_many(filepath, finds, replaces, keep_case, match_word)
        elif process_sub and os.path.isdir(f'{folderpath}{path}/'):
            replace_folder(f'{folderpath}{path}/', finds,
                           replaces, keep_case, match_word, process_sub)


folder = sys.argv[1]
find = sys.argv[2]
replace = sys.argv[3]
keep_case = sys.argv[4] == "true"
match_word = sys.argv[5] == "true"
process_sub = sys.argv[6] == "true"

replace_folder(folder, find, replace, keep_case, match_word, process_sub)
