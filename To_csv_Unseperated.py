# -*- coding: utf-8 -*-
import sys
import re
import importlib
import os
import pandas as pd
import win32com.client
importlib.reload(sys)
import docx

from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from io import StringIO

replacements = {
        "▪":" ", 
        "":" ", 
        "–":"-",
        "":" ",
        "’":"'",
        "·":" ",
        "●":" ",
        "•":" ",
        "“":"'",
        "”":"'",
        "\n":" ",
        "\r":" ",
        "\xa0":" ",
        "\xc0":" "
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
    
    combined_text = " ".join(fullTxt)
    
    rep = dict((re.escape(k), v) for k, v in replacements.items())
    pattern = re.compile("|".join(rep.keys()))
    
    my_str = pattern.sub(lambda m: rep[re.escape(m.group(0))], combined_text)
    
     
    return my_str


def docx_converter(path):
    
    Doc = docx.Document(path)
    
    fullText = []
    
    for paragraph in Doc.paragraphs:
        fullText.append(paragraph.text)
        
    rep = dict((re.escape(k), v) for k, v in replacements.items())
    
    pattern = re.compile("|".join(rep.keys()))
    
    my_str = pattern.sub(lambda m: rep[re.escape(m.group(0))], ' '.join(fullText))
    
    return my_str


def doc_converter(path):
    
    fullText = []
    
    app = win32com.client.Dispatch('Word.Application')
    app.Visible = False 
    app.Documents.Open(path)

    doc = app.ActiveDocument
  
    fullText.append(doc.Content.Text)
    
    doc.Close()
    app.Quit()
    
    rep = dict((re.escape(k), v) for k, v in replacements.items())
    
    pattern = re.compile("|".join(rep.keys()))
    
    my_str = pattern.sub(lambda m: rep[re.escape(m.group(0))], ' '.join(fullText))
    
    return my_str


def main():
    
    folder_path = "C:\\Users\\raymond\\Desktop\\1"
    partial_path = "C:\\Users\\raymond\\Desktop\\1\\"
    df = pd.DataFrame(columns=["Filename", "Content", "Received_interview"])
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
            if file_extension == ".pdf1":
                
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
                output = doc_converter(file_path)
                
#                new_path = doc_to_docx(file_path)
#                
#                output = docx_converter(new_path)
                
                df = df.append(
                            {
                                "Filename":filename, 
                                "Content":"'" + output + "'",
                                "Received_interview":str(subfolder)[0]
                            }, 
                                ignore_index=True
                    )
            
                
    df.to_csv("C:\\Users\\raymond\\Desktop\\labeled_resumes.csv", index=False)
    
    return
            
            
if __name__=='__main__':
    main()
                        