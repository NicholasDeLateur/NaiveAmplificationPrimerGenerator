# written Nicholas A DeLateur in Weiss Lab, MIT; delateur@mit.edu
# Version 1.1
# Run using PyCharm Community Edition 2016.1.3 with Anaconda3 on Windows 10

import xlrd
import xlsxwriter
import datetime
import os
from itertools import takewhile

def column_len(sheet, index):
    col_values = sheet.col_values(index)
    col_len = len(col_values)
    for _ in takewhile(lambda x: not x, reversed(col_values)):
        col_len -= 1
    return col_len

def RC_DNA(sequence):
    # Taken from https://gist.github.com/crazyhottommy/7255638#file-reverse_complement-py
    seq_dict = {'A':'T','T':'A','G':'C','C':'G'}
    return "".join([seq_dict[base] for base in reversed(sequence)])

def CalculateTm(sequence):
    Tm = 0
    for s in sequence:
        if s == 'A' or s == 'T':
            Tm += 2
        else:
            Tm += 4
    return Tm


def CreatePrimers(sequence):
    primer5to3 = ''
    primer3to5 = ''
    sequenceRC = RC_DNA(sequence)
    Tm = 0
    i = 0
    while (Tm < 68):
        primer5to3 += sequence[i]
        Tm = CalculateTm(primer5to3)
        i += 1
        if i >= 40 or i == len(sequence):
            break
    Tm = 0
    i = 0
    while (Tm < 68):
        primer3to5 += sequenceRC[i]
        Tm = CalculateTm(primer3to5)
        i += 1
        if i >= 40 or i == len(sequence):
            break
    primers = [primer5to3,primer3to5]
    return primers


Sequences_of_Interest = []
Primers = []

f = os.path.normpath("C:/Users/Weiss Lab/Desktop/a.xlsx")
workbookRead = xlrd.open_workbook(f)
worksheetRead = workbookRead.sheet_by_index(0)

for i in range(1, column_len(worksheetRead,0)):
    Sequences_of_Interest.append(worksheetRead.row_values(i))
print(Sequences_of_Interest)

for i in range (len(Sequences_of_Interest)):
    Primers.append(CreatePrimers(str(Sequences_of_Interest[i][0])))
print(Primers)
# Create an new Excel file and add a worksheet.
# WARNING doesn't check if the filename already exists.
# Shouldn't be a problem if you don't run the program within seconds of itself
# (ROW, COLUMN)....(LEFT, TOP)
filename = 'AmplicationPrimers' + datetime.datetime.now().strftime("%y%m%d%H%M%S") + '.xlsx'
workbook1 = xlsxwriter.Workbook(filename)
worksheet1 = workbook1.add_worksheet()
worksheet1.write(0, 1, 'Name')
worksheet1.write(0, 2, '5to3 primer')
worksheet1.write(0, 3, 'Tm')
worksheet1.write(0, 4, '3to5 primer')
worksheet1.write(0, 5, 'Tm')
worksheet1.write(0, 0, 'Sequence')

for i in range(len(Sequences_of_Interest)):
    worksheet1.write(i + 1, 0, Sequences_of_Interest[i][0])
    worksheet1.write(i + 1, 1, Sequences_of_Interest[i][1])
    worksheet1.write(i + 1, 2, Primers[i][0])
    worksheet1.write(i + 1, 3, CalculateTm(Primers[i][0]))
    worksheet1.write(i + 1, 4, Primers[i][1])
    worksheet1.write(i + 1, 5, CalculateTm(Primers[i][1]))

# MANDATORY close the workbook
workbook1.close()