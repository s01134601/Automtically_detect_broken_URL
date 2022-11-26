#

import os
import requests
from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
import pandas as pd


def gen_doc_list(source) -> list:
    """
    returns relevant files from source directory
    :param source: path to directory with relevant files
    :return: list of relevant files
    """
    doc_list = []
    for root, dirnames, filenames in os.walk(source):
        for filename in filenames:
            if filename.endswith(('.doc', '.docx')):
                doc_list.append(os.path.join(root, filename))
    return doc_list


def extract_url_from_doc(doc_path):
    doc_url_df = pd.DataFrame()
    path = []
    url = []
    document = Document(doc_path)
    rels = document.part.rels
    # print(rels)
    for rel in rels:
        if rels[rel].reltype == RT.HYPERLINK:
            if rels[rel]._target.startswith('http'):
                path.append(doc_path)
                # doc_dict[rel] = rels[rel]._target
                url.append(rels[rel]._target)
                # print(rels[rel]._target)
    doc_url_df['doc_path'] = path
    doc_url_df['url'] = url
    return doc_url_df


def get_url_status(doc_url_df):
    stat_codes = []
    for index, row in doc_url_df.iterrows():
        try:
            response = requests.get(row.url, verify=False, timeout=10.0)
            stat_codes.append(response.status_code)
        except requests.exceptions.Timeout:
            stat_codes.append(504)
        except requests.exceptions.ConnectionError:
            stat_codes.append(-1)
    doc_url_df['status'] = stat_codes
    return doc_url_df


if '__main__' == __name__:

    # find doc and docx files
    requests.packages.urllib3.disable_warnings()
    source_dir = '/Users/bling/Desktop/test_folder'
    doc_list = gen_doc_list(source_dir)
    print(doc_list)

    # find urls inside files
    for doc_path in doc_list:
        doc_url_df = extract_url_from_doc(doc_path)
        doc_url_df = get_url_status(doc_url_df)
    doc = doc_url_df.loc[doc_url_df['status']!= 200]

doc.to_csv('/Users/bling/Desktop/test_folder/broken_links/broken_list_doc.csv')

