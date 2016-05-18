# written Nicholas A DeLateur in Weiss Lab, MIT; delateur@mit.edu
# Version 1.0
# Run using PyCharm Community Edition 2016.1.3 with Anaconda3 on Windows 10

#import xlsxwriter
#import datetime

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


def CreatePrimer(sequence):
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
    print(Tm, ' ', i)
    Tm = 0
    i = 0
    while (Tm < 68):
        primer3to5 += sequenceRC[i]
        Tm = CalculateTm(primer3to5)
        i += 1
        if i >= 40 or i == len(sequence):
            break
    print(Tm, ' ', i)
    primers = [primer5to3,primer3to5]
    return primers


Sequences_of_Interest = ['ATGACAGTAAAGAAGCTTTATTTCGTCCCAGCAGGTCGTTGTATGTTGGATCATTCGTCTGTTAATAGTACATTAACACCAGGAGAATTATTAGACTTACCGGTTTGGTGTTATCTTTTGGAGACTGAAGAAGGACCTATTTTAGTAGATACAGGTATGCCAGAAAGTGCAGTTAATAATGAAGGTCTTTTTAACGGTACATTTGTCGAAGGGCAGGTTTTACCGAAAATGACTGAAGAAGATAGAATCGTGAATATTTTAAAACGGGTTGGTTATGAGCCGGAGGACCTTCTTTATATTATTAGTTCTCACTTGCATTTTGATCATGCAGGAGGAAATGGCGCTTTTATAAATACACCAATCATTGTACAGCGTGCTGAATATGAGGCGGCGCAGCATAGCGAAGAATATTTGAAAGAATGTATATTGCCGAATTTAAACTACAAAATCATTGAAGGTGATTATGAAGTCGTACCAGGAGTTCAATTATTGCATACACCAGGCCATACTCCAGGGCATCAATCGCTATTAATTGAGACAGAAAAATCCGGTCCTGTATTATTAACGATTGATGCATCGTATACGAAAGAGAATTTTGAAAATGAAGTGCCATTTGCGGGATTTGATTCAGAATTAGCTTTATCTTCAATTAAACGTTTAAAAGAAGTGGTGATGAAAGAGAAGCCGATTGTTTTCTTTGGACATGATATAGAGCAGGAAAGGGGATGTAAAGTGTTCCCTGAATATATATAA',
]
Primers = [CreatePrimer(s) for s in Sequences_of_Interest]
print(Primers)


# Create an new Excel file and add a worksheet.
# WARNING doesn't check if the filename already exists.
# Shouldn't be a problem if you don't run the program within seconds of itself
#filename = 'AmplicationPrimers' + datetime.datetime.now().strftime("%y%m%d%H%M%S") + '.xlsx'
#workbook1 = xlsxwriter.Workbook(filename)
#worksheet1 = workbook1.add_worksheet()
#worksheet1.write(0, 0, 'Name')
#worksheet1.write(0, 1, '5to3 primer')
#worksheet1.write(0, 2, 'Tm')
#worksheet1.write(0, 3, '3to5 primer')
#worksheet1.write(0, 4, 'Tm')
#worksheet1.write(0, 5, 'Sequence')
#for i in range(0,M):
#    worksheet1.write(i+1, 0, GeneratedSequences[i])

# MANDATORY close the workbook
#workbook1.close()