import sys
import subprocess
import random
from datetime import date
from datetime import timedelta
import xml.etree.ElementTree as et

try:
    import openpyxl
except ModuleNotFoundError as e:
    print("Openpyxl module missing. Run 'pip3 install openpyxl' to run this program again")
    exit()


def main():
    group = []
    week = []
    nextMonday=str(date.today()+timedelta(days=(7-date.today().weekday())))
    DAYSOFWEEK=('Monday','Tuesday', 'Wednesday', 'Thursday', 'Friday')
    engtree=et.parse("MESS_list.xml")
    engroot=engtree.getroot()

    for child in engroot.findall("Eng"):
        id=child.attrib
        if child[2].text=='RTP' and child[1].text != 'CP':
            #print(id['CEC'])
            group.append((id['CEC'],child[0].text,child[1].text,child[2].text))


    random.shuffle(group)
    for i in range(5):
        week.insert(i,group[:])
        random.shuffle(group)

    wb=openpyxl.Workbook()
    sheet=wb["Sheet"]
    sheet.title=nextMonday
    workday=[]
    r=1
    c=1
    for day, cday in zip(week, DAYSOFWEEK):
        sheet.cell(row=r, column=c, value=cday)
        r+=1
        for person in day:
            workday.append(str(person[1])+"("+str(person[0])+")"+str(person[2]))
        for i in workday:
            sheet.cell(row=r, column=c, value=i)
            r+=1
        r=1
        c+=2
        workday=[]

    wb.save(nextMonday+".xlsx")

if __name__ == '__main__':
    main()

