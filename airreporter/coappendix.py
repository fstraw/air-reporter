# -*- coding: utf-8 -*-
"""
Created on Wed Jun 10 12:09:37 2015

@author: bbatt
"""
import os

import pyfpdf
from PyPDF2 import PdfFileReader, PdfFileMerger

conds = ("Existing.in", "NoBuild.in", "Build.in",
           "Existing.out", "NoBuild.out", "Build.out")

ws = r""

def print_model_file(ws, of):
    """Convert CO model to pdf"""
    f = open(os.path.join(ws, of), 'rb')
    if of.endswith('.in'):
        pdfname = os.path.join(ws, '{}_In.pdf'.format(of.replace('.in','')))
    elif of.endswith('.out'):
        pdfname = os.path.join(ws, '{}_Out.pdf'.format(of.replace('.out','')))
    pdf = pyfpdf.FPDF(format='letter')
    pdf.add_page()
    pdf.set_font('Arial', size=8)
    for line in f.readlines():
        if 'PAGE' in line and not 'PAGE  1' in line:
            pdf.add_page()
        pdf.write(3.5, line)
    pdf.output(pdfname)
    return pdfname

#Specify location of CO inputs/outputs
def create_co_appendix(ws):
    """Return pdf of CO outputs/inputs"""    
    merger = PdfFileMerger()
    pdflist = []
    for of in conds:
        pdflist.append(print_model_file(ws, of))
    for pdf in pdflist:
       merger.append(PdfFileReader(pdf, "rb"))
       os.remove(pdf)
    pdfname = os.path.join(ws, "CO_ALL.pdf")
    merger.write(pdfname)
    return pdfname

if __name__ == '__main__':
    create_co_appendix(ws)