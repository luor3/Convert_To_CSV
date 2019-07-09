# -*- coding: utf-8 -*-

import sys
import docx 
import importlib
import filetype
import os
importlib.reload(sys)

from pdfminer.pdfparser import PDFParser, PDFDocument
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import LTTextBoxHorizontal, LAParams
from pdfminer.pdfinterp import PDFTextExtractionNotAllowed


def pdf_converter(path):
    
    fullText = []
    
    fp = open(path, 'rb')

    praser = PDFParser(fp)
    # create a pdf file
    doc = PDFDocument()
    
    praser.set_document(doc)
    doc.set_parser(praser)
    
    doc.initialize()
    #
    #  check if the file is extractable
    if not doc.is_extractable:
        raise PDFTextExtractionNotAllowed
    #
    # if is not extractable, then ignore it    
    else:
        
        rsrcmgr = PDFResourceManager()
        
        laparams = LAParams()
        device = PDFPageAggregator(rsrcmgr, laparams=laparams)
        
        interpreter = PDFPageInterpreter(rsrcmgr, device)
        #
        # loop through each page
        for page in doc.get_pages():
            interpreter.process_page(page)
            #
            # layout is a LTPage object, it contains like TTextBox, LTFigure, LTImage, LTTextBoxHorizontal
            layout = device.get_result()
            
            for x in layout:
                if(isinstance(x, LTTextBoxHorizontal)):
                    results = x.get_text().replace("▪", "").\
                                           replace("","").\
                                           replace("–","-").\
                                           replace("","")
                                           
                fullText.append(results)
            
    return fullText


def docx_converter(path):
    
    fullText = []
    # open the docx file
    doc = docx.Document(path)
    #
    # read the doc file
    for paragraph in doc.paragraphs:
        
        fullText.append(paragraph.text)

    return fullText


def main():
    
    folder_path = "C:\\Users\\raymond\\Desktop\\8"
    partial_path = "C:\\Users\\raymond\\Desktop\\8\\"
    #
    # read all files in the folder
    for filename in os.listdir(folder_path):
        
        full_path = partial_path + str(filename)
        #
        # if the filename type is pdf
        if filetype.guess_extension(full_path) == "pdf":
            
            pdf_converter(full_path)
        #
        # the file type is docx
        else:
            docx_converter(full_path)
            
            
if __name__=='__main__':
    main()
                        