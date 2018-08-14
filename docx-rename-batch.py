#!/usr/bin/env python

"""
Author: Autoz

Extract title from word:docx file.

Depends on: docx2txt.

Origin from: [pdftitle.py](https://gist.github.com/hanjianwei/6838974)

Usage:
    post-docx-batch.py *.docx
"""

import io # cStringIO in python2
import getopt
import os
import re
import string
import sys
import docx2txt

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

def copyright_line(line):
    """Judge if a line is copyright info.
    """
    return re.search(r'technical\s+report|proceedings|preprint|to\s+appear|submission', line.lower())

def empty_str(s):
    return len(s.strip()) == 0

def title_start_end_postal(lines,search_lines=30,tag=''):
    filted_lines = []; find_tag = False;
    search_lines = search_lines if search_lines > len(lines) else len(lines)
    for i, line in enumerate(lines[:search_lines]):
        if not empty_str(line):
            if check_contain_chinese(line): # user name or chinese title
                if tag in line and check_contain_number(line): # tag and numbers as doc's No.
                    filted_lines.append(line.strip())
                    print('=cut line=%s'% line.strip())
                    find_tag = True
                    break
                filted_lines.append(line.strip())
                print('=cut line=%s'% line.strip())
                continue
    return filted_lines,find_tag

def docx_title(filename,tag):
    text = docx2txt.process(filename)
    lines,find_tag = title_start_end_postal(text,search_lines=30,tag=tag)
    title = ('_'.join(line.strip() for line in lines).replace(' ','_'))
    if not find_tag: title = title[:40]
    if empty_str(title):
        title = os.path.splitext(filename)[0]
        print('Can not get info to rename.')
    new_name = os.path.join('', title + os.path.splitext(filename)[1])
    return new_name

if __name__ == "__main__":

    def usage():
        print("Usage: %s [--dry-run] [-p|--post-tag post_tag] filenames" % sys.argv[0])
        input('There is no flie in root folder, please enter anykey to exit')

    try:
        opts, args = getopt.getopt(sys.argv[1:], 'np:', ['post-tag','dry-run'])
    except getopt.GetoptError:
        usage()

    from glob import glob
    #args = glob('*.docx') # for build in pyinstall to .exe apps with no arguments
    #print('====docs====',args)

    if not args: usage()

    readall = False
    dry_run = False
    post_tag = '编号' # stop tag

    for opt, arg in opts:
        if opt in ['-a', '--readall']:
            readall = True
        elif opt in ['-n', '--dry-run']:
            dry_run = True
        elif opt in ['-p', '--post-tag']:
            post_tag = arg

    for filename in args:
        new_name = docx_title(filename,post_tag)
        print ("%s => %s" % (filename,new_name))
        if not dry_run:
            if(filename == new_name): continue
            if os.path.exists(new_name):
                new_name = os.path.splitext(new_name)[0] + "copy" + os.path.splitext(new_name)[1]
                print("===Target already exists! Rename to %s" % new_name)
            os.rename(filename, new_name)
    input('Done, please enter anykey to exit')
    