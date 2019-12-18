# -*- coding: utf-8 -*-
"""
Created on Fri Oct 25 14:24:38 2019

@author: NMacEwan
"""

from tkinter import Tk
from tkinter import filedialog as dlg

def getfile(initpath, title='Select file', multi=False):
    """ Allow user to select single or multiple Excel Files """

    types = (("Excel Files", "*.xls*"), ("All files", "*.*"))

    if multi:
        Tk().withdraw()
        filename = dlg.askopenfilenames(initialdir=initpath, title=title, filetypes=types)    
    else:
        Tk().withdraw()
        filename = dlg.askopenfilename(initialdir=initpath, title=title, filetypes=types)
        
    return filename


def getfolder(initpath, title):
    """ Allow user to select folder """
    Tk().withdraw()

    folder = dlg.askdirectory(initialdir=initpath, title=title)

    return folder

if __name__ == "__main__":
    test = getfile("C:\\", True)

    for file in test:
        print(file)
        print("-----")
