import tkinter
import openpyxl as op
import numpy

def f():
    "This is excel"
    #print("Let's start")
    wb = op.load_workbook(filename='test.xlsx')
    average_age = 0
    average_score = 0
    sheet = wb['test']
    val = sheet['A1'].value
    #print(val)
    
    ids = []
    age = []
    score = []
  
    for x in sheet['A2':'A4']:
        for y in x:
            ids.append(y.value)
    
    for x in sheet['B2':'B4']:
        for y in x:
            age.append(y.value)
    
    for x in sheet['C2':'C4']:
        for y in x:
            score.append(y.value)
    average_age = av_ag(age)
    average_score = av_ag(score)
    print("Items:",len(ids))
    print("Average age:",average_age)
    print("Max score:",max(score))
    print("Min score:",min(score))
    print("Average score:",average_score)

def av_ag(age):
    a = 0
    for x in age:
        a += x
        
    return round(a/len(age),1)
    
