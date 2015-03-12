import win32com.client
import codecs
import sys
import os

class easyExcel:
    """A utility to make it easier to get at Excel.  Remembering
    to save the data is your problem, as is  error handling.
    Operates on one workbook at a time."""

    def __init__(self, filename=None):
        self.xlApp = win32com.client.Dispatch('Excel.Application')
        if filename:
            self.filename = filename
            self.xlBook = self.xlApp.Workbooks.Open(filename)
        else:
            self.xlBook = self.xlApp.Workbooks.Add()
            self.filename = ''  
    
    def save(self, newfilename=None):
        if newfilename:
            self.filename = newfilename
            self.xlBook.SaveAs(newfilename)
        else:
            self.xlBook.Save()    

    def close(self):
        self.xlBook.Close(SaveChanges=0)
        del self.xlApp
        
    def getCell(self, sheet, row, col):
        "Get value of one cell"
        sht = self.xlBook.Worksheets(sheet)
        return sht.Cells(row, col).Value

    def setCell(self, sheet, row, col, value):
        "set value of one cell"
        sht = self.xlBook.Worksheets(sheet)
        sht.Cells(row, col).Value = value

    def getRange(self, sheet, row1, col1, row2, col2):
        "return a 2d array (i.e. tuple of tuples)"
        sht = self.xlBook.Worksheets(sheet)
        return sht.Range(sht.Cells(row1, col1), sht.Cells(row2, col2)).Value

    def setRange(self, sheet, leftCol, topRow, data):
        """insert a 2d array starting at given location. 
        Works out the size needed for itself"""

        bottomRow = topRow + len(data) - 1
        rightCol = leftCol + len(data[0]) - 1
        sht = self.xlBook.Worksheets(sheet)
        sht.Range(
            sht.Cells(topRow, leftCol), 
            sht.Cells(bottomRow, rightCol)
            ).Value = data

    def getContiguousRange(self, sheet, row, col):
        """Tracks down and across from top left cell until it
        encounters blank cells; returns the non-blank range.
        Looks at first row and column; blanks at bottom or right
        are OK and return None witin the array"""

        sht = self.xlBook.Worksheets(sheet)

        #search "Case ID"
        while self.getCell(sheet,row,1) != "Case ID":
            row = row + 1

        row = row + 1
        # find the bottom row
        bottom = row
        while sht.Cells(bottom + 1, col+1).Value not in [None, '']:
            bottom = bottom + 1

        # right column
        right = col
        while sht.Cells(row, right + 1).Value not in [None, '']:
            right = right + 1

        return sht.Range(sht.Cells(row, col), sht.Cells(bottom, right)).Value

    def fixStringsAndDates(self, aMatrix):
        # converts all unicode strings and times
        newmatrix = []
        for row in aMatrix:
            newrow = []
            for cell in row:
                if type(cell) is UnicodeType:
                    newrow.append(str(cell))
                elif type(cell) is TimeType:
                    newrow.append(int(cell))
                else:
                    newrow.append(cell)
            newmatrix.append(tuple(newrow))
        return newmatrix

def usage():
    print "python translatorexcel2xml.py excel_file_name sheet_name suite_name"
    sys.exit("Missing arguments")

cwd = os.getcwd()
sys.path.append(cwd)

if __name__ == "__main__":

    if len(sys.argv) < 3 :
        usage()
    
    excel_file = os.path.join(cwd,sys.argv[1])
    sheet_name = sys.argv[2]
    suite = sys.argv[3]
    
    xml_file = os.path.join(cwd,"%s.xml"%suite)
    
    excel = easyExcel(excel_file)
    xml = codecs.open(xml_file,encoding='utf-8',mode='w+')
    xml.write('''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n''')
     
    arrays = excel.getContiguousRange(sheet_name,1,1)
    xml.write('''<testsuite name="%s">\n'''%suite)
    row = 1
    for array in arrays:
        if array[0] in [None,'']:
            if row > 1 :
                xml.write('''</testsuite>\n''')
            xml.write('''<testsuite name="%s">\n'''%array[1])
        else:
            case = array[1].replace("&", "&amp;")
            case = case.replace("<", "&lt;")
            case = case.replace(">", "&gt;")
            case = case.replace('"', "&quot;")
            
            xml.write('''<testcase name="%s %s">\n'''%(array[0],case))
            xml.write('''<summary>%s</summary>\n'''%case)
            xml.write('''</testcase>\n''')
        row = row + 1
    xml.write('''</testsuite>\n''')
    xml.write('''</testsuite>\n''')
    excel.close()
    xml.close()


