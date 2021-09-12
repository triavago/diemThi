import statistics

from selenium import webdriver
from time import sleep
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
import openpyxl
import csv
#get data from excel file
wb = openpyxl.load_workbook('G:/hoc/DANH SÁCH  XẾP LỚP 10 TẠM THỜI  NĂM HỌC 2021-2022.xlsx')
sheet = wb['Table 1']
SbdList = []
for i in range(5, 283):
    readVal = sheet['B' + str(i)].value
    if readVal == None:
        continue
    if readVal == 'SBD':
        SbdList.append(readVal)
        continue
    if readVal / 10000 < 1:
        readVal = '00' + str(readVal)
    elif readVal / 100000 < 1:
        readVal = '0' + str(readVal)
    SbdList.append(str(readVal))
#print(SbdList)
#print(len(SbdList))
subString = "Tổng điểm XT: "
#open edge
browser = webdriver.Edge(executable_path="msedgedriver.exe")
#access to link
browser.get("http://xemdiem.hanoimoi.com.vn/")
cBoxLop10 = Select(browser.find_element_by_id("type"))
cBoxLop10.select_by_value('1')

gradeList = []
f = open('G:/project/diemThi/venv/diemThi.csv', 'w')

classList = [[] for _ in range(6)]

txtSbd = browser.find_element_by_id("hoten")
for i in SbdList:
    if i == 'SBD':
        gradeList.append(i)
        continue
    txtSbd.send_keys(i)
    txtSbd.send_keys(Keys.ENTER)
    sleep(1)
    fullGrade = browser.find_element_by_xpath("/html/body/div/div[5]/table/tbody/tr/td[4]")
    subStringIndex = fullGrade.text.find(subString)
    totalGrade = fullGrade.text[(subStringIndex + len(subString)): (subStringIndex + len(subString)) + 4]
    gradeList.append(totalGrade)
    txtSbd.clear()

for classes in classList:
    for grade in gradeList:
        if grade == 'SBD':
            gradeList = gradeList[gradeList.index('SBD') + 1:]
            break
        else:
            classes.append(float(grade))
c = 1

for classes in classList:
    print(classes)
    print (statistics.mean(classes))
    print(statistics.median(classes))
    #print("trung binh A" + c + " :" + statistics.mean(classes) + '\n')
    #print("trung vi: " + 'A' + str(c) + statistics.median(classes) + '\n')
    c += 1
#print(classList)

#sleep(5)

#browser.close()
