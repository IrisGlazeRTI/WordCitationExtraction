# This is a sample Python script.

# Press ⌃R to execute it or replace it with your code.
# Press Double ⇧ to search everywhere for classes, files, tool windows, actions, and settings.

# Author. (). Title. _Journal_. Publish_date_month_year; Volume(Issue): Pages.

import os
import pandas as pd
import docx
import docx2txt
import subprocess
import re
subprocess.call('dir', shell=True)


def print_hi(name):
    # Use a breakpoint in the code line below to debug your script.
    print(f'Hi, {name}')  # Press ⌘F8 to toggle the breakpoint.

def readtxt(filename):
    doc = docx.Document(filename)
    fullText = []
    for para in doc.paragraphs:
        fullText.append(para.text)
    return '\n'.join(fullText)

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    # FILE PATH AND creating LIST of all documents
    filepath = r"./input"
    files = os.listdir(filepath)
    document_list = []
    cnames = ['p_id', 'p_name']
    for name in files:
        x = name
        # For each file we find,we need to ensure it is a .docx file before adding it to our list
        if os.path.splitext(os.path.join('/', name))[1] == ".docx":
            if not name.startswith('~$'):
                document_list.append(os.path.join(filepath, name))
    file_path = document_list[0]

    worddoc_to_text = docx2txt.process(file_path)
    split_worddoc_text_arr = worddoc_to_text.split('\n\n')
    doc_title_text = split_worddoc_text_arr.pop(0)
    doc_date_text = split_worddoc_text_arr.pop(0)
    citations_arr = split_worddoc_text_arr
    citations_data_dict = {'author': [], 'title': []}
    for citation in citations_arr:
        author = citation.split('. (', 2)[0]
        remaining_text = '(' + citation.split('. (', 2)[1]
        title_text = remaining_text.split('.')[1]
        citations_data_dict['author'].append(author)
        citations_data_dict['title'].append(title_text)
        print(author, title_text)
    citations_data_dict['journal'] = []
    citations_data_dict['publish_date'] = []
    citations_data_dict['volume_issue_pages'] = []
    citations_data_dict['volume'] = []
    citations_data_dict['issue'] = []
    citations_data_dict['pages'] = []


    doc = docx.Document(file_path)
    idx = 0
    for para in doc.paragraphs:
        if idx >= 2:
            print(para.text)
            runs_len = para.runs.__len__()
            journal = ''
            text_after_journal = ''
            run_idx = 0
            isItalic = False
            startedItalic = False
            endedItalic = False
            section_after_journal = ''
            while run_idx < runs_len and endedItalic == False:
                isItalic = (para.runs[run_idx].italic == True)
                if isItalic:
                    journal += para.runs[run_idx].text.lstrip('.')
                    startedItalic = True
                    x = journal
                elif startedItalic and not endedItalic:
                    endedItalic = True
                    text_after_journal = para.runs[run_idx].text
                run_idx += 1
            citations_data_dict['journal'].append(journal)
            while run_idx < runs_len:
                text_after_journal += para.runs[run_idx].text
                run_idx += 1
            split_last_text_arr = text_after_journal.split(';')
            publish_date = split_last_text_arr[0] if split_last_text_arr.__len__() > 0 else ''
            publish_date = publish_date.strip('.').strip(' ')
            citations_data_dict['publish_date'].append(publish_date)
            volume_issue_pages = split_last_text_arr[1] if split_last_text_arr.__len__() > 1 else ''
            volume_issue_pages = volume_issue_pages.strip('.').strip(' ')
            citations_data_dict['volume_issue_pages'].append(volume_issue_pages)
            volume = ''
            issue = ''
            pages = ''
            #t = "20(3)"
            t = volume_issue_pages
            m = re.match("(\d+)([^:]*)(:?\D*[0-9\-]*)", t)
            if m:
                volume = m.group(1)
                issue = m.group(2).strip().lstrip('(').rstrip(')')
                pages = m.group(3).lstrip(':').strip()
            citations_data_dict['volume'].append(volume)
            citations_data_dict['issue'].append(issue)
            citations_data_dict['pages'].append(pages)
        idx += 1

    print(readtxt(file_path))

    citations_data_df = pd.DataFrame(citations_data_dict)

    html = citations_data_df.to_html()
    f = open('output.html', 'w+')
    f.write(html)
    print(html)
    print_hi('PyCharm')

