#!/usr/bin/env python
    # shebang line; has to be 1st line in python script; for operating system (tells OS that all the following code is python)
# -*- coding: utf-8 -*-

"""
Given a directory path, search all files in the path for a given text string
within the 'word/document.xml' section of a MSWord .dotm file.
"""
__author__ = "jhoelzer"

# how to get into a directory and find something in it

import os
import zipfile
import sys
import argparse # intermediary between command line and code (?)

def create_parser():
    parser = argparse.ArgumentParser(description='searches dotm files for text') # parser object; can now add options in form of arguments
    parser.add_argument('--dir', help='directory to search', default='.') # -- makes something optional; help is a hint as to what the argument does; default is the default directory to search through, the . is the current directory
    parser.add_argument('text', help='text to search for')
    return parser
    # this all just creates it, not runs it


def main(directory, text_to_search): # directory to look in and what to look for
    print('I will search all dotm files in {}'.format(directory))
    print('I will search for magic text {}'.format(text_to_search))
    file_list = os.listdir(directory) # file_list is a list of all the files (in this case, in the dotm_files folder)
    files_searched = 0
    files_matched = 0
    print('Searching directory {} for text {}'.format(directory, text_to_search))
    # for loop to look through each file
    for file in file_list:
        files_searched += 1
        full_path = os.path.join(directory, file)
        if not file.endswith('.dotm'): # checks to see if file is a dotm
            # print('{} is not a dotm file'.format(file))
            continue
        if not zipfile.is_zipfile(full_path): # uses a builtin method from the imported zipfile to check if the file is a zip
            # print('{} is not a zip file'.format(full_path))
            continue
        with zipfile.ZipFile(full_path, 'r') as zipped: # looks inside zip files and tries to read them
            # zip files have compressed data; content starts with 'PK'; can store files within files
            # once uncompressed, 
            table_of_contents = zipped.namelist() # table of contents for each file
            # print(table_of_contents)
            if 'word/document.xml' in table_of_contents:
                with zipped.open('word/document.xml', 'r') as doc:
                    for line in doc:
                        i = line.find(text_to_search) # search for specific thing in text
                        if i >= 0:
                            files_matched += 1
                            # print(i) # prints index of found thing
                            print('{}: ...{}...'.format(file, line[i - 40:i + 40]))
    print('Files searched: {}'.format(files_searched))
    print('Files matched: {}'.format(files_matched))



if __name__ == '__main__':
    parser = create_parser() # this runs the created parser
    ns = parser.parse_args() # ns = namespace = a dictionary accessed with dot notation
    # if ns.dir is not None: # they gave an explicit directory to search
    # ns.text
    # main('dotm_files', '$') # search the dotm_files for '$'
    main(ns.dir, ns.text)
