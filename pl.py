import tkinter
import openpyxl as op
import numpy
import io

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
    mas1 = ["Items","Average age","Max score","Min score:","Average score"]
    mas2 = []
    mas2.append(len(ids))
    mas2.append(average_age)
    mas2.append(max(score))
    mas2.append(min(score))
    mas2.append(average_score)
    
    print("Items:",len(ids))
    print("Average age:",average_age)
    print("Max score:",max(score))
    print("Min score:",min(score))
    print("Average score:",average_score)
    

    with open("res.txt",'w') as file:
        i = 0
        for x in mas1:
            file.write(mas1[i])
            file.write(' - ')
            file.write(str(mas2[i]))
            file.write('\n')
            i += 1
            
        #i = 0
        #for x in file:
        #    x.writelines(mas1[i])
        #    i += 1
    

def av_ag(age):
    a = 0
    for x in age:
        a += x
        
    return round(a/len(age),1)
    
