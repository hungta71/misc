from PyPDF2 import PdfReader
import zipfile
import re
import os, time, sys
from pathlib import Path
import pandas
from colorama import init
from colorama.ansi import Fore
from colorama.initialise import deinit
from send2trash import send2trash
import csv
import docx2txt

init(autoreset=True)

attachment_store_path = 'E:\\OIE\\Attachment_Part 2\\'
noted_store_path = 'E:\\_Docs\\'
standby_store_path = 'E:\\_Standby2\\'
extension_list = ['docx', 'xls', 'xlsx', 'pdf', 'txt', 'log', 'csv']
list_of_keywords = []

with open('E:\\keywords.csv', 'r', encoding="utf-8") as keyword_files: # get keyword list
    reader = csv.reader(keyword_files)
    for row in list(reader):
        if row[1] != '':
            list_of_keywords.append(row[1])
        if row[5] != '':
            list_of_keywords.append(row[5])

class file_extractor:
    def __init__(self, extension = '', modified_time = '', number_of_pages = -1, keyword_found = None, code = 0) -> None:
        self.extension = extension
        self.modified_time = modified_time
        self.number_of_pages = number_of_pages
        self.keyword_found = keyword_found
        self.code = code
    
    def getExtension(self):
        return self.extension

def find_keyword_in_text(source, keyword_list):
    for k in keyword_list:
        if k in source:
            return k
    return None

def pdf_extractor(file_to_check, keyword_list):
    result = {'number_of_pages': -1, 'keyword_found': None, 'modified_time': '', 'error_code': 0}
    try:
        result['modified_time'] = time.ctime(os.path.getmtime(file_to_check))
        f = open(file_to_check, 'rb')
        pdf = PdfReader(f)
        result['number_of_pages'] = len(pdf.pages)
        if result['number_of_pages'] >= 10:
            result['error_code'] = 2
        
        # for i in range(0, result['number_of_pages']):
        #     page_content = pdf.pages[i].extract_text()
        #     result['keyword_found'] = find_keyword_in_text(page_content, keyword_list)
        #     if result['keyword_found'] is not None:
        #         break
    except Exception as e:
        print(Fore.LIGHTRED_EX + "Error opening file: " + Fore.LIGHTYELLOW_EX + file_to_check + Fore.RESET + ' - ' + str(e))
        result['error_code'] = 1
    f.close()
    return result

def docx_extractor(file_to_check, keyword_list):
    result = {'number_of_pages': -1, 'keyword_found': None, 'modified_time': '', 'error_code': 0}
    try:
        result['modified_time'] = time.ctime(os.path.getmtime(file_to_check))
        text = docx2txt.process(file_to_check)
        result['keyword_found'] = find_keyword_in_text(text, list_of_keywords)
    except Exception as e:
        print(Fore.LIGHTRED_EX + "Error opening file: " + Fore.LIGHTYELLOW_EX + file_to_check + Fore.RESET + ' - ' + str(e))
        result['error_code'] = 1
    return result

def doc_extractor(file_to_check, keyword_list):
    result = {'number_of_pages': -1, 'keyword_found': None, 'modified_time': '', 'error_code': 0}
    try:
        result['modified_time'] = time.ctime(os.path.getmtime(file_to_check))
        docx_file = file_to_check + 'x'
        if not os.path.exists(docx_file):
            os.system('antiword "' + file_to_check + '" > "' + docx_file + '"')
            # with open(docx_file) as f:
            #     text = f.read()
            # result['keyword_found'] = find_keyword_in_text(text, keyword_list)
            # os.remove(docx_file)
    except Exception as e:
        print(Fore.LIGHTRED_EX + "Error opening file: " + Fore.LIGHTYELLOW_EX + file_to_check + Fore.RESET + ' - ' + str(e))
    return result

def excel_extractor(file_to_check, keyword_list):
    result = {'number_of_pages': -1, 'keyword_found': None, 'modified_time': '', 'error_code': 0}
    try:
        result['modified_time'] = time.ctime(os.path.getmtime(file_to_check))
        df = pandas.read_excel(file_to_check)
        text = df.to_string()
        result['keyword_found'] = find_keyword_in_text(text, list_of_keywords)
    except Exception as e:
        print(Fore.LIGHTRED_EX + "Error opening file: " + Fore.LIGHTYELLOW_EX + file_to_check + Fore.RESET + ' - ' + str(e))
        result['error_code'] = 1
    return result

def powerpoint_extractor(file_to_check, keyword_list):
    result = {'number_of_pages': -1, 'keyword_found': None, 'modified_time': '', 'error_code': 0}
    
    try:
        result['modified_time'] = time.ctime(os.path.getmtime(file_to_check))

        archive = zipfile.ZipFile(file_to_check, "r")
        ms_data = archive.read("docProps/app.xml")
        archive.close()
        app_xml = ms_data.decode("utf-8")
        regex = r"<(Pages|Slides)>(\d)</(Pages|Slides)>"
        matches = re.findall(regex, app_xml, re.MULTILINE)
        match = matches[0] if matches[0:] else [0, 0]
        result['number_of_pages'] = match[1]

        result['keyword_found'] = any(map(app_xml.__contains__, keyword_list))
    except Exception as e:
        print(Fore.LIGHTRED_EX + "Error opening file: " + Fore.LIGHTYELLOW_EX + file_to_check + Fore.RESET + ' - ' + str(e))
        result['error_code'] = 1

    return result

def general_extractor(file_to_check, keyword_list):
    result = {'number_of_pages': -1, 'keyword_found': None, 'modified_time': '', 'error_code': 0}
    with open(file_to_check, 'r') as f:
        try:
            result['modified_time'] = time.ctime(os.path.getmtime(file_to_check))
            text = f.read()
            result['keyword_found'] = find_keyword_in_text(text, list_of_keywords)
        except Exception as e:
                print(Fore.LIGHTRED_EX + 'Error opening file ' + Fore.LIGHTYELLOW_EX + file_to_check + Fore.RESET + ' - ' + str(e))
    return result

def file_extractor(file_to_check, keyword_list, file_extension):
    if file_extension == 'pdf':
        return pdf_extractor(file_to_check, keyword_list)
    elif file_extension == 'docx':
        return docx_extractor(file_to_check, keyword_list)
    elif file_extension == 'doc':
        return doc_extractor(file_to_check, keyword_list)
    elif file_extension == 'pptx':
        return powerpoint_extractor(file_to_check, keyword_list)
    elif file_extension == 'xlsx' or file_extension == 'xls':
        return excel_extractor(file_to_check, keyword_list)
    else:
        return general_extractor(file_to_check, keyword_list)

def process_file(path_to_folder):
    i,j = 0,0
    file_list = Path(path_to_folder).glob('*\\*.*')
    for file in file_list:
        file_path = str(file)
        temp = file_path.split('\\')
        file_name = temp[-1]
        temp = file_name.split('.')
        file_extension = temp[-1].lower()

        if file_extension in extension_list:
            i += 1
            check_result = file_extractor(file_to_check=file_path, keyword_list=list_of_keywords, file_extension=file_extension)
            if check_result['keyword_found'] is not None:
                print('Found: ' + Fore.LIGHTGREEN_EX + file_path + Fore.RESET + ' - pages: ' + Fore.LIGHTGREEN_EX + str(check_result['number_of_pages']) + Fore.RESET + ' - Keyword: ' + Fore.GREEN + check_result['keyword_found'])
                if check_result['number_of_pages'] >= 4:
                    os.rename(file_path, noted_store_path + file_name)
                j += 1
            elif check_result['error_code'] == 0:
                send2trash(file_path)
            elif check_result['error_code'] == 2:
                print("Move to standby: " + Fore.LIGHTMAGENTA_EX + file_path + Fore.RESET + ' - pages: ' + Fore.LIGHTMAGENTA_EX + str(check_result['number_of_pages']))
                os.rename(file_path, standby_store_path + file_name)

    print("Total file processed: " + Fore.YELLOW + str(i) + Fore.RESET + " - Total file found: " + Fore.GREEN + str(j))

if __name__ == '__main__':
    start = time.time() 
    process_file(attachment_store_path)
    deinit()
    end = time.time()
    print('Elapsed time: ' + str(round(end - start, 2)) + 's')