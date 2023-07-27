# This is a sample Python script.

# Press ⌃R to execute it or replace it with your code.
# Press Double ⇧ to search everywhere for classes, files, tool windows, actions, and settings.

# Author. (). Title. _Journal_. Publish_date_month_year; Volume(Issue): Pages.

import os
import pandas as pd
import docx
import subprocess
import re
import docxpy
from datetime import datetime
import unicodecsv as csv
#from docx import Document
#from docx.opc.constants import RELATIONSHIP_TYPE as RT

subprocess.call('dir', shell=True)


def output_file(file_path=''):
    input_file_name = file_path.split('/')[file_path.split('/').__len__() - 1] if file_path.split('/').__len__() > 0 else ''
    output_directory_name = 'output'
    if not os.path.exists('./' + output_directory_name):
        os.makedirs('./' + output_directory_name)
    output_file_path = './' + output_directory_name + '/' + input_file_name.rstrip('.xdoc')
    docxpy_doc = docxpy.DOCReader(file_path)
    docxpy_doc.process()  # process file
    split_worddoc_text_arr = docxpy_doc.data['document'].strip('\n').split('\n\n')

    doc_title_text = split_worddoc_text_arr.pop(0)
    doc_date_text = split_worddoc_text_arr.pop(0)
    citations_arr = split_worddoc_text_arr
    citations_data_dict = {'author': [], 'publications': [], 'title': []}
    for citation in citations_arr:
        author = citation.split('. (', 2)[0]
        publications_year_1 = citation.split('. (', 2)[1].split(')')[0] if citation.split('. (', 2).__len__() >= 2 and citation.split('. (', 2)[1].split(')').__len__() > 0 else ''
        remaining_text = '(' + citation.split('. (', 2)[1]
        title_text = remaining_text.split('.')[1]
        citations_data_dict['author'].append(author)
        citations_data_dict['publications'].append(publications_year_1)
        citations_data_dict['title'].append(title_text)
        print(author, title_text)
    citations_data_dict['publish_date'] = []
    # blank col
    citations_data_dict['description'] = []
    citations_data_dict['journal'] = []
    #citations_data_dict['volume_issue_pages'] = []
    citations_data_dict['volume'] = []
    citations_data_dict['issue'] = []
    citations_data_dict['pages'] = []
    # blank col
    citations_data_dict['doi'] = []
    citations_data_dict['pm_ids'] = []
    citations_data_dict['links'] = []
    #blank cols
    citations_data_dict['epub'] = []
    citations_data_dict['created'] = []
    citations_data_dict['updated'] = []
    citations_data_dict['pub_type'] = []
    citations_data_dict['date_year'] = []
    citations_data_dict['protocol_ids'] = []
    citations_data_dict['study_name'] = []
    citations_data_dict['study_acronym'] = []
    citations_data_dict['study_type'] = []
    citations_data_dict['disease_phenotype'] = []
    citations_data_dict['primary_research_focus'] = []
    citations_data_dict['funding_source'] = []
    citations_data_dict['award_num'] = []
    citations_data_dict['foa'] = []

    pubmed_base_url = 'https://pubmed.ncbi.nlm.nih.gov/'

    doc = docx.Document(file_path)
    idx = 0
    for para in doc.paragraphs:
        if idx >= 2:
            first_para_hyperlink_id = para.paragraph_format.element.xpath('./w:hyperlink')[0].attrib[
                '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id']
            link_url = para.part.rels[first_para_hyperlink_id].target_ref
            pm_id = None
            if link_url.startswith(pubmed_base_url):
                pm_id = link_url.lstrip(pubmed_base_url).split('/')[0] \
                    if link_url.lstrip(pubmed_base_url).split('/').__len__() > 0 else ''
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
            # journal is defined now
            while run_idx < runs_len:
                text_after_journal += para.runs[run_idx].text
                run_idx += 1
            split_last_text_arr = text_after_journal.split(';')
            publish_date = ''
            pattern_year_month = re.compile("^([0-9]{4}) ([A-Za-z]+)$")
            pattern_year_only = re.compile("^([0-9]{4})$")

            if split_last_text_arr.__len__() > 0:
                '''
                x = re.compile("([0-9]{4}) ([A-Za-z]+)").match(split_last_text_arr[0])
                z = split_last_text_arr[0][28:]
                y = pattern_year_only.match(split_last_text_arr[0])
                '''
                publish_date = split_last_text_arr[0]
                publish_date = publish_date.strip('.').strip(' ')
                publish_date = publish_date[re.compile("([0-9]{4}) ?([A-Za-z]+)?").search(publish_date).regs[0][0]:re.compile("([0-9]{4}) ?([A-Za-z]+)?").search(publish_date).regs[0][1]]

            is_date_year_month = bool(pattern_year_month.match(publish_date))
            is_date_year_only = bool(pattern_year_only.match(publish_date))
            publish_datetime_object = None
            if is_date_year_month:
                publish_datetime_object = datetime.strptime(publish_date, '%Y %b')
            elif is_date_year_only:
                publish_datetime_object = datetime.strptime(publish_date, '%Y')

            publish_date_reformatted = publish_datetime_object.strftime('%Y-%m-%d %H:%M:%S') if publish_datetime_object is not None else ''
            # publish_date_reformatted is defined now.
            volume_issue_pages = split_last_text_arr[1] if split_last_text_arr.__len__() > 1 else ''
            volume_issue_pages = volume_issue_pages.strip('.').strip(' ')
            #citations_data_dict['volume_issue_pages'].append(volume_issue_pages)
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
            citations_data_dict['publish_date'].append(publish_date_reformatted)
            citations_data_dict['journal'].append(journal)
            citations_data_dict['volume'].append(volume)
            citations_data_dict['issue'].append(issue)
            citations_data_dict['pages'].append(pages)
            citations_data_dict['pm_ids'].append(pm_id)
            citations_data_dict['links'].append(link_url)
            # blank cols
            citations_data_dict['description'].append('')
            citations_data_dict['doi'].append('')
            citations_data_dict['epub'].append('')
            citations_data_dict['created'].append('')
            citations_data_dict['updated'].append('')
            citations_data_dict['pub_type'].append('citing')
            citations_data_dict['date_year'].append('')
            citations_data_dict['protocol_ids'].append('')
            citations_data_dict['study_name'].append('')
            citations_data_dict['study_acronym'].append('')
            citations_data_dict['study_type'].append('')
            citations_data_dict['disease_phenotype'].append('')
            citations_data_dict['primary_research_focus'].append('')
            citations_data_dict['funding_source'].append('')
            citations_data_dict['award_num'].append('')
            citations_data_dict['foa'].append('')

        idx += 1

    citations_data_df = pd.DataFrame(citations_data_dict)

    csv = citations_data_df.to_csv()
    html = citations_data_df.to_html()

    f = open(output_file_path + '.html', 'w+', encoding='utf-8')
    f.write(html)
    print(html)

    f = open(output_file_path + '.csv', 'w+', encoding='utf-8')
    f.write(csv)

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

    for document in document_list:
        output_file(document)



