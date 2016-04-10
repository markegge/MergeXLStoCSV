
# coding: utf-8

# In[1]:

import xlrd
import glob
import csv
import re


# In[2]:

def gen_opener(filenames):
    '''
    Open a sequence of filenames one at a time producing a file object.
    The file is closed immediately when proceeding to the next iteration.
    '''
    for filename in filenames:
        if filename.endswith('.gz'):
            f = gzip.open(filename, 'rt')
        elif filename.endswith('.bz2'):
            f = bz2.open(filename, 'rt')
        else:
            f = open(filename, 'rt')
        yield f
        f.close()

def gen_concatenate(iterators):
    '''
    Chain a sequence of iterators together into a single sequence.
    '''
    for it in iterators:
        for rator in it:
            yield rator


# In[3]:

def gen_write_out():
    with open("combined.csv", 'wb') as f:
        writer = csv.writer(f, quoting=csv.QUOTE_MINIMAL)
        while True:
            try:
                output = yield
                writer.writerow(output)
            except GeneratorExit:
                print 'closed file'
                f.close()


# In[7]:

writer = gen_write_out()
writer.send(None)
header_needed = True
file_names = glob.glob("*.xls")   

for filename in file_names:
    print filename
    book = xlrd.open_workbook(filename, on_demand=True)
    print book.sheet_names()
    for i in range(book.nsheets):
        sheet = book.sheet_by_index(i)
        print "sheet:", i, sheet.name
        for row in xrange(sheet.nrows):
            rowValues = sheet.row_values(row)
            # going to parse for route numbers here...
            #routes = ['191', '444', '900', '722', '614']
            routes = [444, 900, 611, 612, 613, 614, 711, 712, 713, 714] # filter to just G2, P1, 61x, and 71x...
            newValues = []
            try:
                if int(float(rowValues[2])) in routes:
                    for s in rowValues:
                        if isinstance(s, unicode):
                            strValue = (str(s.encode("utf-8")))
                        else:
                            strValue = (str(s))

                        isInt = bool(re.match("^([0-9]+)\.0$", strValue))

                        if isInt:
                            strValue = int(float(strValue))
                        else:
                            isFloat = bool(re.match("^([0-9]+)\.([0-9]+)$", strValue))
                            isLong  = bool(re.match("^([0-9]+)\.([0-9]+)e\+([0-9]+)$", strValue))

                            if isFloat:
                                strValue = float(strValue)

                            if isLong:
                                strValue = int(float(strValue))

                        newValues.append(strValue)
                    writer.send(newValues)
            except ValueError as c: # if header...
                if header_needed:
                    writer.send(rowValues)
                    header_needed = False
                    
        book.unload_sheet(i)
    book.release_resources()
try:
    writer.close()
except GeneratorExit:
    print "Shutting down."

