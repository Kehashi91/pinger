import re
import openpyxl
from openpyxl.utils import get_column_letter

xlsfile = 'testsmall.xlsx'

def checkmail(cell):
    mailre = re.compile(r'([^\s]+)@([\w]+)\.([\w]+)')
    if cell and mailre.match(cell.strip()):
        return True
    else:
        return False
        
def checkurl(cell):
    checkurlre = re.compile(r'zzz/[0-9]{6}')#!!!!!!!!!!!!!
    if cell and checkurlre.match(cell.strip()):
        return True
    else:
        return False

def rowverification(row):
    desc = ""
    link = ""
    for cell in row:
        if checkmail(cell.value):
            print("leeeel juz jest mail") 
            pass
        elif checkurl(cell.value) and link:
            raise ValueError("W wierszu {} istnieją zduplikowane linki".format(cell.row))
        elif checkurl(cell.value) and not link:
            print("leeeel to url") 
            link = cell.value
        elif cell.value and desc:
            raise ValueError("W wierszu {} istnieją zduplikowane opisy".format(cell.row))
        elif cell.value:
            print("leeeel to desc") 
            desc = cell.value
    if desc and link:        
        return (desc, link)
    elif not desc:
        raise ValueError("W wierszu {} brakuje opisu.".format(row[0].row))
    elif not link:
        raise ValueError("W wierszu {} brakuje linku.".format(row[0].row))
            
def readxlsfile(xlsfile):
    workbook = openpyxl.load_workbook(xlsfile)
    sheet = workbook.get_sheet_names()[0]
    activesheet = workbook.get_sheet_by_name(sheet)
    return activesheet
    #return activesheet['A1':'{0}{1}'.format(maxcol, maxrow)]

def checkemptyrows(worksheet, position):
        rowtocheck = [cell.value for cell in worksheet[position - 1] if cell.value != None]
        if rowtocheck:
            return True
        elif not rowtocheck:
            return False
    
def tablemap(worksheet):
    mailmapfinal  = []
    mailmap = []
    mailadresses = []
    for column in tuple(worksheet.columns):
        for cell in column:
            if checkmail(cell.value):
                mailmap.append(cell.row)
                mailadresses.append(cell.value)
    mailmap.append(worksheet.max_row)
    for mail in mailmap[0:len(mailmap) - 1]:
        if mailmap.index(mail) == len(mailmap) - 2:
            print("Tru")
            mailmapfinal.append((mail, mailmap[mailmap.index(mail) + 1]))
        elif checkemptyrows(worksheet, mail):
            print("Tru1")
            mailmapfinal.append((mail, mailmap[mailmap.index(mail) + 1] - 1))
        elif not checkemptyrows(worksheet, mail):
            mailmapfinal.append((mail, mailmap[mailmap.index(mail) + 1] - 2))
            print("Tru2")
    print(mailmapfinal)
    print(mailadresses)
    messagedataconstructor(worksheet, mailmapfinal, mailadresses[0])
        
def messagedataconstructor(worksheet, mailmap, mailadress):
    messagedict = {}
    for object in mailmap:
        responselist = []
        for row in worksheet[object[0]:object[1]]:
            responselist.append(rowverification(row))
        messagedict[mailadress] = responselist
    print (messagedict)

if __name__ == "__main__":
    worksheet = readxlsfile(xlsfile)
    tablemap(worksheet)
    
    
"""

for row in data:
    message = []
    response = []
    rowprint = [cell.value for cell in row if cell.value != None]
    for cell in rowprint:
        link = ""
        if checkmail(cell):
            message.append(cell)
        elif checkhyperlink(cell):
            links.append(cell)
        elif cell:
            response.append(tup(link, response))
    print(rowprint)
print(message)
print(links)
"""
"""
test = ("   keke@beyond.pl",)
test2 = ("",)
test3 = ("aaaabababab dlsksdk kroolick.",)
test4 = ("dziaba dziaba. kolmp@dies.pl",)
checkmail(test)
checkmail(test2)
checkmail(test3)
checkmail(test4)

print(message)
"""
  