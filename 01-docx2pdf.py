#!/usr/bin/env python3
# -*- encoding: utf-8 -*-
'''
@File    :   01-docx2pdf.py
@Time    :   2024/08/12 15:32:37
@Author  :   Wanjin.Hu 
@Version :   1.0
@Contact :   wanjin.hu@outlook.com
@Description : docx 批量转换为 pdf
'''

from docx2pdf import convert
import os
import argparse
import time

parser = argparse.ArgumentParser(description="docx to pdf")
parser.add_argument("-d", "--project_dir", dest="projectDir", type=str, required=True, help="Input project dir")
args = parser.parse_args()
def docx_to_pdf(docx_path, pdf_path):
    convert(docx_path, pdf_path)

def batch_convert_docx_to_pdf(folder_path):
    for filename in os.listdir(folder_path):
        if filename.endswith('.docx'):
            docx_file = os.path.join(folder_path, filename)
            pdf_file = os.path.join(folder_path, filename[:-5] + '.pdf')
            docx_to_pdf(docx_file, pdf_file)
            print(f'Converted {filename} to {pdf_file}')
            time.sleep(1)

def main():
    folder_path = args.projectDir
    batch_convert_docx_to_pdf(folder_path)
    print('All .docx files converted to .pdf successfully!')

if __name__ == '__main__':
    main()
