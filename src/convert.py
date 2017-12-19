# author: Vigneshwar Padmanaban
# date created: Dec 09, 2017

from sys import argv
import openpyxl
import sys
import shutil

import os
from openpyxl.styles import Color, PatternFill, Font, Border
# from openpyxl.styles import colors
# from openpyxl.cell import Cell
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.chart import (
    ScatterChart,
    Reference,
    Series,
)
import pyexcel as p

class To_xlsx:
    @staticmethod
    def clean():
        for the_file in os.listdir('uploads'):
            file_path = os.path.join('uploads', the_file)
            try:
                if os.path.isfile(file_path):
                    os.unlink(file_path)
                    # elif os.path.isdir(file_path): shutil.rmtree(file_path)
            except Exception as e:
                print(e)
    @staticmethod
    def convert(inputExcel):
        To_xlsx.clean()
        fileName = str(inputExcel)
        splitName = inputExcel.split('.')
        splitName[1] = inputExcel[-3:]
        if(splitName[1] == 'xls'):
            print('Its a xls file. Converting to .xlsx file..')
            p.save_book_as(file_name=fileName,dest_file_name=splitName[0]+'.xlsx')
            opFile = splitName[0]+'.xlsx'
            destpath = 'uploads'
            srcpath = opFile
            shutil.move(srcpath, destpath)
            print('File moved!')
            return opFile
        elif(splitName[1]=='lsx'):
            print('Its already a xlsx file')
            opFile = inputExcel
            destpath = 'uploads'
            srcpath = opFile
            shutil.move(srcpath, destpath)
            return opFile
        else:
            print('Invalid file format! Upload .xls files')
            sys.exit()
