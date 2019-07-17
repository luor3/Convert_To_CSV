# -*- coding: utf-8 -*-
import sys
import re
import docx 
import importlib
import os
import pandas as pd
import win32com.client
importlib.reload(sys)


from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from io import StringIO

replacements = {
        "▪":"", 
        "":"", 
        "–":"-",
        "":"",
        "’":"'",
        "·":"",
        "●":"",
        "•":"",
        "“":"'",
        "”":"'",
        "é":"e",
        "\n":"",
        "\r":""
        }

def pdf_to_txt(path):
    fullTxt = []
    
    rsrcmgr = PDFResourceManager()
    
    retstr = StringIO()
    
    codec = 'utf-8'
    
    laparams = LAParams()
    
    device = TextConverter(rsrcmgr, retstr, codec=codec, laparams=laparams)
    
    fp = open(path, 'rb')
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    
    password = ""
    maxpages = 0
    caching = True
    pagenos=set()

    for page in PDFPage.get_pages(
                                  fp, 
                                  pagenos, 
                                  maxpages=maxpages,     
                                  password=password,
                                  caching=caching, 
                                  check_extractable=True
    ):
        
        interpreter.process_page(page)

    fullTxt.append(retstr.getvalue()
    )
    

    fp.close()
    device.close()
    retstr.close()
    
    combined = " ".join(fullTxt)
    
    rep = dict((re.escape(k), v) for k, v in replacements.items())
    pattern = re.compile("|".join(rep.keys()))
    
    my_str = pattern.sub(lambda m: rep[re.escape(m.group(0))], combined)
    
     
    return my_str


def docx_converter(path):
    
    fullText = []
    # open the docx file
    doc = docx.Document(path)
    
    section = doc.sections[0]
    
    header = section.header
    
    for paragraph in header.paragraphs:
        fullText.append(paragraph.text)
    #
    # read the doc file
    for paragraph in doc.paragraphs:
        
        fullText.append(paragraph.text)
    
    combined = " ".join(fullText)
    
    rep = dict((re.escape(k), v) for k, v in replacements.items())
    pattern = re.compile("|".join(rep.keys()))
    
    my_str = pattern.sub(lambda m: rep[re.escape(m.group(0))], combined)
    

    return my_str


def doc_to_docx(path):
    
    w = win32com.client.Dispatch('Word.Application')
    
    w.Visible = 0
    w.DisplayAlerts = 0
    
    doc = w.Documents.Open(path)
    
    newpath = os.path.splitext(path)[0] + '.docx'
    
    doc.SaveAs(newpath, 12, False, "", True, "", False, False, False, False)
    
    doc.Close()
    w.Quit()
    os.remove(path)
    
    return newpath


def main():
    
    folder_path = "C:\\Users\\raymond\\Desktop\\resumes"
    partial_path = "C:\\Users\\raymond\\Desktop\\resumes\\"
    df = pd.DataFrame(columns = ["Filename", "Content", "Received_interview"])
    #
    # read all subfolder in the folder
    for subfolder in os.listdir(folder_path):
        
        full_path = partial_path + str(subfolder)
        #
        # read all files in subfolder
        for filename in os.listdir(full_path):
            
            file_path = full_path + "\\" +str(filename)
            #
            # identify the file type
            file_extension = os.path.splitext(file_path)[1]
            #
            # if the filename type is pdf
            if file_extension == ".pdf":
                
                output = pdf_to_txt(file_path)               
                # subfolder will be named as 0Big4_noInterview, subfolder[0] will return
                # either 0 or 1
                df = df.append(
                            {
                                "Filename":filename, 
                                "Content":"'" + output + "'",
                                "Received_interview":str(subfolder)[0]
                            }, 
                                ignore_index=True
                )
        
            #
            # the file type is docx
            elif file_extension == ".docx":
                
                docx_converter(file_path)
                
                output = docx_converter(file_path)
                
                df = df.append(
                            {
                                "Filename":filename, 
                                "Content":"'" + output  + "'",
                                "Received_interview":str(subfolder)[0]
                            }, 
                                ignore_index=True
                    )
                
            elif file_extension == ".doc":
                
                new_path = doc_to_docx(file_path)
                
                output = docx_converter(new_path)
                
                df = df.append(
                            {
                                "Filename":filename, 
                                "Content":"'" + output + "'",
                                "Received_interview":str(subfolder)[0]
                            }, 
                                ignore_index=True
                    )
            
                
    df.to_csv("C:\\Users\\raymond\\Desktop\\labelled_resumes.csv", index=False)
    
    return
            
            
if __name__=='__main__':
    main()
                        