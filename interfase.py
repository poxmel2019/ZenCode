from tkinter import *
import openpyxl as op

def f():
    global mas1, mas2
    def handle():
        global mas1, mas2
        wb = op.load_workbook(filename=ent.get())
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
    
        #print("Items:",len(ids))
        #print("Average age:",average_age)
        #print("Max score:",max(score))
        #print("Min score:",min(score))
        #print("Average score:",average_score)
    

        with open("res.txt",'w') as file:
            i = 0
            for x in range(0,len(mas1)):
                file.write(mas1[i])
                file.write(' - ')
                file.write(str(mas2[i]))
                file.write('\n')
                i += 1
            
   
        l['text'] = ent.get()
    def av_ag(age):
        a = 0
        for x in age:
            a += x
        
        return round(a/len(age),1)
    root = Tk()
    Label(text='Введите сюда адрес файла').pack()
    ent = Entry()
    but = Button(text='handle',command=handle)
    ent.pack()
    but.pack()
    l = Label(width=10, height=1)
    l.pack()
    root.mainloop()
