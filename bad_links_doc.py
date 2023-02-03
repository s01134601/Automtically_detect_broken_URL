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
    try:
        document = Document(doc_path)
    except Exception:
        return
    rels = document.part.rels
    for rel in rels:
        if rels[rel].reltype == RT.HYPERLINK:
            if rels[rel]._target.startswith('http'):
                path.append(doc_path)
                url.append(rels[rel]._target)
                # print(rels[rel]._target)
    doc_url_df['doc_path'] = path
    doc_url_df['url'] = url
    return doc_url_df


def get_url_status(doc_url_df_total):
    stat_codes = []
    for index, row in doc_url_df_total.iterrows():
        try:
            response = requests.get(row.url, verify=False, timeout=10.0, allow_redirects= False)
            stat_codes.append(response.status_code)
        except requests.exceptions.Timeout:
            stat_codes.append(504)
        except requests.exceptions.ConnectionError:
            stat_codes.append(-1)
    doc_url_df_total['status'] = stat_codes
    return doc_url_df_total


# if '__main__' == __name__:

def detect_url_doc(input_dir_path, output_folder_path):
    # find doc and docx files
    requests.packages.urllib3.disable_warnings()
    doc_list = gen_doc_list(input_dir_path)
    # find urls inside files
    doc_url_df_total = pd.DataFrame()
    for doc_path in doc_list:
        doc_url_df = extract_url_from_doc(doc_path)
        doc_url_df_total = pd.concat([doc_url_df_total, doc_url_df])
    doc_url_df_total = get_url_status(doc_url_df_total)
    # print(doc_url_df_total)
    doc = doc_url_df_total.drop_duplicates()
    doc = doc.loc[doc['status'] != 200]
    doc.to_csv(os.path.join(output_folder_path, "broken_links_docx.csv"), index=False)

