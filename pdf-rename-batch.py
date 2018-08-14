#!/usr/bin/env python

"""
Extract title from PDF file.

Depends on: pyPDF, PDFMiner.

Usage:

    find . -name "*.pdf" |  xargs -I{} pdf-title-rename-batch.py -d tmp --rename {}
"""

import io # cStringIO in python2
import getopt
import os
import re
import string
import sys

from PyPDF2 import PdfFileReader
from PyPDF2.utils import PdfReadError

from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfinterp import PDFResourceManager, process_pdf, PDFTextExtractionNotAllowed
from pdfminer.pdfparser import PDFSyntaxError

__all__ = ['pdf_title']

def check_contain_chinese(check_str):
    return any((u'\u4e00' <= char <= u'\u9fff') for char in check_str)

def check_contain_number(check_str):
    return any(char.isdigit() for char in check_str)

def sanitize(filename):
    """Turn string to valid file name.
    """
    valid_chars = "-_.() %s%s" % (string.ascii_letters, string.digits)
    return ''.join([c for c in filename if c in valid_chars])
    
def sanitize_chinese(filename):
    return re.sub('\?|\.|\。|\!|\/|\;|\:|\*|\>|\<|\~|\(|\)|\[|\]|[A-Za-z0-9]|', '', filename)

def meta_title(filename):
    """Title from pdf metadata.
    """
    try:
        fp = open(filename, 'rb')
        docinfo = PdfFileReader(fp).getDocumentInfo()
        fp.close()
        print('===docinfo===',docinfo.title)
        return docinfo.title if docinfo.title else ""
    except PdfReadError:
        print(">>>>>Can't not read doc meta info")
        return ""

def copyright_line(line):
    """Judge if a line is copyright info.
    """
    return re.search(r'technical\s+report|proceedings|preprint|to\s+appear|submission', line.lower())

def empty_str(s):
    return len(s.strip()) == 0

def pdf_text(filename):
    try:
        text = io.StringIO()
        rsrc = PDFResourceManager()
        device = TextConverter(rsrc, text, laparams=LAParams()) # no codec='utf-8' in TextConverter
        fp = open(filename, 'rb')
        process_pdf(rsrc, device, fp, None, maxpages=1, password='') # open(filename, 'rb') need to close or use: 'with open(filename, 'rb') as fp'
        #fp.close()
        device.close()
        print('===text.getvalue===',text.getvalue())
        return text.getvalue()
    except (PDFSyntaxError, PDFTextExtractionNotAllowed, UnicodeEncodeError):
        print(">>>>>Can't read doc's text info")
        return ""

def title_start(lines):
    for i, line in enumerate(lines):
        if not empty_str(line) and not copyright_line(line):
            return i;
    return 0

def title_end(lines, start, max_lines=2):
    for i, line in enumerate(lines[start+1:start+max_lines+1], start+1):
        if empty_str(line):
            return i
    return start + 1

def title_start_end(lines, max_lines=100,line_range=4):
    start = max_lines - line_range; end = max_lines # last 4 line
    for i, line in enumerate(lines[:max_lines]):
        if check_contain_chinese(line): # Chinese
            if (" " in line) or check_contain_number(line):
                continue
            start = i; end = i + line_range; break
        elif not empty_str(line) and not copyright_line(line):
            if check_contain_number(line):
                continue
            start = i; end = i + line_range - 1; break
            continue # Other language
    print('=start=%s==end=%s'%(start,end))
    return start,end

def text_title(filename):
    """Extract title from PDF's text.
    """
    lines = pdf_text(filename).strip().split('\n')
    #i = title_start(lines)
    #j = title_end(lines, i,max_lines=5)
    i,j = title_start_end(lines,max_lines=20,line_range=5)
    print('====return===',' '.join(line.strip() for line in lines[i:j]))
    return ' '.join(line.strip() for line in lines[i:j])

def valid_title(title):
    return not empty_str(title) and empty_str(os.path.splitext(title)[1])

def pdf_title(filename):
    title = meta_title(filename)
    if valid_title(title):
        return title
    title = text_title(filename)
    if valid_title(title):
        return title
    return os.path.basename(os.path.splitext(filename)[0])

if __name__ == "__main__":
    opts, args = getopt.getopt(sys.argv[1:], 'nd:', ['dry-run', 'rename'])

    dry_run = False
    rename = True
    dir = "."

    for opt, arg in opts:
        if opt in ['-n', '--dry-run']:
            dry_run = True
            rename = False
        elif opt in ['--rename']:
            rename = True
        elif opt in ['-d']:
            dir = arg

    if len(args) == 0:
        print ("Usage: %s [-d output] [--dry-run] [--rename] filenames" % sys.argv[0])
        sys.exit(1)

    for filename in args:
        title = pdf_title(filename)
        if rename:
            if check_contain_chinese(title):
                title = sanitize_chinese('_'.join(title.split())) # for Chinese
            else:
                title = sanitize(' '.join(title.split())) # for others languages
            new_name = os.path.join('', title + ".pdf") 
            print ("%s => %s" % (filename, new_name))
            if not dry_run:
                if os.path.exists(new_name):
                    print ("*** Target %s already exists! ***" % new_name)
                else:
                    os.rename(filename, new_name)
        else:
            print (title)