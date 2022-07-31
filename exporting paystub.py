import pdfplumber, openpyxl as xl, os, re
from tkinter import filedialog
from datetime import date, timedelta

filepath = os.path.normpath(filedialog.askopenfilename())

# filepath = "C:\\Users\\hoang\\Desktop\\CENERGY\\PAYSTUBS\\07-14-2022.pdf"
with pdfplumber.open(filepath) as pdf:
    page1 = pdf.pages[0]
    words = page1.extract_words()

text_words = []
for i in range(len(words)):
    text_words.append(words[i]["text"])

# print(text_words)

ValueList = []  # start with date

txt = text_words[6].split("/")  # finds first instance of ##/##/#### and splits it *hardcoded position 6
txt = txt[0] + "/" + txt[1] + "/" + txt[2]
ValueList.append(txt)

# print(txt)


#  Adding keywords

TermList = ["Regular", "Overtime", "W/H(S)", "Security", "Medicare", "BASE", "Vision", "Net", "Gross"]
TermListPos = []

for x in TermList:
    if x in text_words:
        TermListPos.append(text_words.index(x))
    else:
        TermListPos.append(0)

# print(TermListPos)

#  Adding keyword values

for x in range(len(TermListPos)):
    if TermListPos[x] == 0:
        ValueList.append('0')
        ValueList.append('0')
        if x < 3:
            ValueList.append('0')
    if x <= 1 and TermListPos[x] != 0:  # refers to Reg/OT, finds hours, pay rate, and net pay
        for y in range(TermListPos[x] + 1, TermListPos[x] + 4):  # appends the next 3 items
            ValueList.append(text_words[y])
    if 2 <= x <= 6 and TermListPos[x] != 0:  # finds taxes (fed, ss, and medicare) and deductions (dental and vision)
        for y in range(TermListPos[x] + 1, TermListPos[x] + 3):  # appends next 2 items
            ValueList.append(text_words[y])
    if x == 7:  # refers to net pay, finds net pay
        ValueList.append(text_words[TermListPos[x] + 2])  # Net Amt
    if x == 8:  # refers to gross pay, finds gross pay
        ValueList.append(text_words[TermListPos[x] + 2])  # Gross Amt
        ValueList.append(text_words[TermListPos[x] + 3])  # Gross YTD

# print(ValueList)

# Converting to float values

ValueListInt = []
for x in ValueList:
    if '/' in x:
        ValueListInt.append(x)
    elif '$' in x:
        x = x.split('$')[1]
        if ',' in x:
            x = x.split(',')
            x = x[0] + x[1]
            ValueListInt.append(float(x))
        else:
            ValueListInt.append(float(x))
    elif ',' in x:
        pass
    else:
        ValueListInt.append(float(x))

# print(ValueListInt)

# Week Management

start_date = date(2022, 6, 2)

txt = txt.split("/")
(yy, mm, dd) = (int(txt[2]), int(txt[0]), int(txt[1]))
weekly_date = date(yy, mm, dd)

delta = (weekly_date - start_date) / timedelta(days=7) / 2  # divide by 2 bc paid biweekly

print(delta)

# Excel exporting
workbook = xl.load_workbook(filename="spl pay.xlsx")
sheets = workbook.sheetnames
sheet = workbook["SPL"]

col_num = 2 + delta  # starting from second column


def excel():
    for row in range(1, 21):
        try:
            sheet.cell(row, col_num).value = ValueListInt[row-1]
        except IndexError:
            pass


excel()
workbook.save(filename="spl pay.xlsx")
os.startfile("spl pay.xlsx")


