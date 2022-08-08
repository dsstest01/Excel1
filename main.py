# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
# See PyCharm help at https://www.jetbrains.com/help/pycharm/
import openpyxl as xl
import os,sys,xlsx

#Main Program
Infile = 'transactions.xlsx'
Outfile= 'transactions2.xlsx'
xlsx.updateWb(Infile,Outfile)

os.chdir(sys.path[0])
os.startfile(Outfile)


