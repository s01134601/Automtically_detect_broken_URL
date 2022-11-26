import pandas as pd
import PyPDF2
import os
import requests

def gen_pdf_list(source) -> list:
    pdf_list = []
    for root, dirnames, filenames in os.walk(source):
        for filename in filenames:
            if filename.endswith('.pdf'):
                pdf_list.append(os.path.join(root, filename))
    return pdf_list


def get_url_status(pdf_df):
    stat_codes = []
    for index, row in pdf_df.iterrows():
        try:
            response = requests.get(row.url, verify=False, timeout=10.0)
            stat_codes.append(response.status_code)
        except requests.exceptions.Timeout:
            stat_codes.append(504)
        except requests.exceptions.ConnectionError:
            stat_codes.append(-1)
    pdf_df['status'] = stat_codes
    return pdf_df



pdf_url_df = pd.DataFrame()
source_dir = '/Users/bling/Desktop/test_folder'
pdf_list = gen_pdf_list(source_dir)
path_list = []
url_list = []
page_n = []
pdf_df = pd.DataFrame()
# Open The File in the Command
for pdfs in pdf_list:
    PDFFile = open(pdfs, 'rb')
    PDF = PyPDF2.PdfFileReader(PDFFile)
    pages = PDF.getNumPages()
    key = '/Annots'
    uri = '/URI'
    ank = '/A'

    for page in range(pages):
        page_df = pd.DataFrame(columns=['url', 'page_n', "path"])
#        print("Current Page: {}".format(page))
        pageSliced = PDF.getPage(page)
        pageObject = pageSliced.getObject()
        if key in pageObject.keys():
            ann = pageObject[key]
            for a in ann:
                u = a.getObject()
                if uri in u[ank].keys():
                    if not u[ank][uri].startswith('mailto'):
                        url_list.append(u[ank][uri])
                        path_list.append(pdfs)
                        page_n.append(page+1)

pdf_df['path'] = path_list
pdf_df['url'] = url_list
pdf_df['page_n'] = page_n

requests.packages.urllib3.disable_warnings()
pdf_df = get_url_status(pdf_df)
pdf_df.to_csv('/Users/bling/Desktop/test_folder/broken_links/broken_list_pdf.csv')


