import os
import xlrd
import csv
import codecs
import jellyfish
import fuzzywuzzy

def match(value1, value2):
    maximum = float(max(len(value1), len(value2)))
    distance = jellyfish.damerau_levenshtein_distance(value1, value2)
    return 1 - distance / maximum

def xlsx_to_csv(name_of_excel):
    workbook = xlrd.open_workbook('%s.xlsx'% name_of_excel)
    table = workbook.sheet_by_index(0)
    with codecs.open('%s.csv' % name_of_excel, 'w', encoding='utf-8') as f:
        write = csv.writer(f)
        for row_num in range(table.nrows):
            row_value = table.row_values(row_num)
            write.writerow(row_value)

if __name__ == '__main__':
    xlsx_to_csv(r'C:\Users\347898222\Desktop\编程\excel项目\stata')
    xlsx_to_csv(r'C:\Users\347898222\Desktop\编程\excel项目\银行名单')

process.extractOne(S1,ListS)