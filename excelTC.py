import re
import sys
import shutil
import yaml 
import os

sys.path.append ('C:\\Python27\\Lib\\xlrd-0.9.3')
sys.path.append ('C:\\Python27\\Lib\\xlutils-1.7.1')
sys.path.append ('C:\\Python27\\Lib\\xlwt-0.7.5')

import xlrd
import xlutils
import xlwt
from xlutils.copy import copy


def main():
    if len(sys.argv) != 2 or sys.argv[1] == '-h':
        sys.exit("excelTC.py <TC file>")

    stream = open (sys.argv[1], 'r') 
    testCase = yaml.load (stream) 
    print testCase 

    readWB = xlrd.open_workbook('C:\\JE-Python\\convertexcel\\RevisedTestCase_format4.xls')
    readSheet = readWB.sheet_by_name('TestCase')
    titleRow = readSheet.row_values (0)

    writeWB = copy (readWB)
    writeSheet = writeWB.get_sheet(0)

    style = xlwt.XFStyle()
    style.alignment.wrap = 1

    # handle single line values
    for i, elem in enumerate(titleRow): 
        if re.search ("Steps", elem):
            print "SKIP " + elem
        else:
            print str (i) + " : " + elem + " : " +  str (testCase[elem])
            writeSheet.write (1, i, testCase[elem], style)

    # handle multiple line values, ie, Design Steps
    steps = testCase["Design Steps"]
    for i, step in enumerate (steps):
#        print "\n" + str(step) + "\n"
        writeSheet.write (1 + i, 14, step["Step Name"], style)
        writeSheet.write (1 + i, 15, step["Description"], style)
        writeSheet.write (1 + i, 16, step["Expected"], style)

             
    newFile = os.path.splitext(sys.argv[1])[0] + ".xls"
    try:
        os.remove (newFile)
    except:
        pass
           
    writeWB.save (newFile) 

main()
