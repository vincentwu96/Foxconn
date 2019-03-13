import PyPDF2, io
import urllib.request as urllib2
import datetime, sys
import openpyxl
import os
from xlrd import open_workbook,cellname





# Parsing PDF
# URL = input("Enter URL\n")
URL = "file:///C:/Users/vincentwu/Desktop/Parse/RegularTicket.pdf"
req = urllib2.Request(URL, headers={'User-Agent' : "Magic Browser"})
remote_file = urllib2.urlopen(req).read()
memory_file = io.BytesIO(remote_file)

read_pdf = PyPDF2.PdfFileReader(memory_file)
number_of_pages = read_pdf.getNumPages()

text = ""
for i in range(0, number_of_pages):
    pageObj = read_pdf.getPage(i)
    page = pageObj.extractText()
    text += page


print(text)
# Formatting PDF Output
format1 = text.split("kok.yee@foxconn.com", 1)[1]   # Delete front elements
format1 = format1.rsplit("Description", 1)[0]       # Delete back elements
format2 = format1.split("\n")
output = ""                                         # Description and Date
Partnum = ""
Serialnum = ""
Vendor = ""
counter = 1        # Counter for Description and Date
counter2 = 1       # Counter for Partnum, Serialnum, Vendor
flag = False
for x in range(0,len(format2)):
    if(format2[x] == "Vendor"):
        Partnum += str(counter2) + ") " + format2[x+1] + "\n"
        Serialnum += str(counter2) + ") " + format2[x+2] + "\n"
        Vendor += str(counter2) + ") " + format2[x+3] + "\n"
        int(counter2)
        counter2+=1
    if(format2[x] == "kok.yee@foxconn.com"):
        if(flag == True):
            break
        y = x
        z = y+4
        output += str(counter) + ") " + format2[x+1] + " "
        int(counter)
        counter+=1
        for y in range(y,len(format2)):
            
            if(format2[y] == "Description"):
                break
            elif(format2[y] == "Vendor"):
                z+=11
            elif(format2[y] == "INITIAL"):
                last = len(format2) - 1 - format2[::-1].index('kok.yee@foxconn.com')
                output = ' '.join(output.split(' ')[:-1])
                output = ' '.join(output.split(' ')[:-1])
                output += format2[last+1] + " "
                for i in range(last+5, len(format2)):
                    output += format2[i] + "\n"
                flag = True
                break
            else:
                if(y>z):                    
                    output += format2[y] + "\n"
                    
output = output.rstrip()                    
print(output)

Partnum = Partnum.rstrip()
Serialnum = Serialnum.rstrip()
Vendor = Vendor.rstrip()
print(Partnum)
print(Serialnum)
print(Vendor)

#os.system("pause")

