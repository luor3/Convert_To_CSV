# -*- coding: utf-8 -*-
import re
import os
import pandas as pd
import win32com.client
import docx
import PyPDF2

replacements = {
    "▪" : " ", 
    "" : " ", 
    "–" : "-",
    "" : " ",
    "’" : "'",
    "·" : " ",
    "●" : " ",
    "•" : " ",
    "“" : "'",
    "”" : "'",
    "\n" : " ",
    "\r" : " ",
    "\xa0" : " ",
    "\xc0" : " "
}

def pdf_to_txt(path):

    fullText = []
    file = open(path, "rb") 
    fileReader = PyPDF2.PdfFileReader(file)
    
    for pageNum in range(fileReader.numPages):
        pageObj = fileReader.getPage(pageNum)
        fullText.append(pageObj.extractText())
    
    paragraphs = replacement(fullText)
    
    return paragraphs


def docx_converter(path):
    
    doc = docx.Document(path)
    
    fullText = []
    
    for paragraph in doc.paragraphs:
        fullText.append(paragraph.text)
        
    paragraphs = replacement(fullText)
    
    return paragraphs


def doc_converter(path):
    
    fullText = []
    
    app = win32com.client.Dispatch('Word.Application')
    app.Visible = False 
    app.Documents.Open(path)

    doc = app.ActiveDocument
  
    fullText.append(doc.Content.Text)
    
    doc.Close()
    app.Quit()
    
    paragraphs = replacement(fullText)
    
    return paragraphs


def replacement(fullText):
    
    rep = dict((re.escape(k), v) for k, v in replacements.items())  
    pattern = re.compile("|".join(rep.keys()))
    
    content = ' '.join(fullText)  
    paragraphs = pattern.sub(lambda m: rep[re.escape(m.group(0))], content)
    
    return paragraphs


def append_to_df(filename, subfolder, output, df):
    # subfolder will be named as 0Big4_noInterview, subfolder[0] will return
    # either 0 or 1
    df = df.append(
        {
            "Filename" : filename, 
            "Content" : "'" + output + "'",
            "Received_interview" : str(subfolder)[0]
            }, 
            ignore_index=True
        )
    
    return df
    

def main():
    
    base_path = os.path.dirname(__file__)
    folder_path = os.path.join(base_path, "1")
    partial_path = os.path.normcase(folder_path) + "\\"
    
    df = pd.DataFrame(columns=["Filename", "Content", "Received_interview"])
    
    for subfolder in os.listdir(folder_path):
        
        full_path = partial_path + str(subfolder)
      
        for filename in os.listdir(full_path):
            
            file_path = full_path + "\\" + str(filename)
           
            file_extension = os.path.splitext(file_path)[1]
           
            if file_extension == ".pdf":
                
                output = pdf_to_txt(file_path)  
                df = append_to_df(filename, subfolder, output, df)
                
            elif file_extension == ".docx":
               
                output = docx_converter(file_path)
                df = append_to_df(filename, subfolder, output, df)
                
            elif file_extension == ".doc":
                
                output = doc_converter(file_path)
                df = append_to_df(filename, subfolder, output, df)
            
                
    df.to_csv(os.path.join(base_path,"labeled_resume.csv"), index=False)
    
    return
            
            
if __name__=='__main__':
    main()
                        