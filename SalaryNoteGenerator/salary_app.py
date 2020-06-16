import tkinter as tk
import openpyxl
import tkinter.filedialog
import salary_note_generator as GM
import os
import time

def openFile():
    def progress():
        tk.Lable(window, text='提取进度：',).place(x=50, y=60)
        canvas = tk.Canvas(window, width=300, height=22, bg="white")
        canvas.place(x=110, y=60)
        fill_line = canvas.create_rectangle(1.5, 1.5, 0, 23, width=0, fill="green")

    def operate(sheet_name):
        ws = wb[sheet_name]
        GM.extract_gui(ws)
        print('\n\n完成！\n\n')

    def operate_new(sheet_name):
        note = '%s 提取进度'%sheet_name
        label = tk.Label(window, text=note,)
        label.place(x=15, y=250)
        canvas = tk.Canvas(window, width=300, height=22, bg="white")
        canvas.place(x=150, y=250)
        fill_line = canvas.create_rectangle(1.5, 1.5, 0, 23, width=0, fill="green")
        ws = wb[sheet_name]
        salary_info = [[i.value for i in each] for each in ws]
        title = salary_info[0]
        persons = salary_info[1:]
        count = len(persons)
        n = 300/count # 300是矩形填充满的次数
        row = 1
        for each in persons:
            try:
                n += 300/count
                name = each[1] + '_' + each[0] + '.xlsx'
                row += 1
            except TypeError as e:
                row += 1
                with open('logs.txt','a',encoding='utf-8') as f:
                    info = str(each[1])+' _ '+ str(each[0])
                    location = ' 第%s行 ##'% row
                    content = f.split('\\')[-1] + '##' + location + info + ':' + str(e)+'\n'
                    f.write(content)
                continue
            data = GM.vertical_data(title, each)
            GM.produce_book(name, data)
            canvas.coords(fill_line, (0, 0, n, 250))
            window.update()

        time.sleep(1)
        label.destroy()
        canvas.destroy()

    def getSheet(event):
        choice = LB.get(LB.curselection())
        # print(type(choice))
        sheet = choice.split('"')[1]
        print(choice)
        print('\n开始处理表单...\n')
        operate_new(sheet)
        LB.destroy()

    f = tk.filedialog.askopenfilename(filetypes=[('Excel文件','.xlsx')])
    wb = openpyxl.open(f)
    sheets = wb.worksheets
    LB = tk.Listbox(window, width=50)
    a = LB.bind('<Double-Button-1>',getSheet)
    for i in sheets:
        LB.insert(tk.END, i)
    LB.pack()


window = tk.Tk()
window.title('个人工资信息提取APP')
window.geometry('500x300')
choice = tk.Button(window,text='打开工作簿', command=openFile)
choice.pack()
if not os.path.exists('个人工资信息'):
    os.mkdir('个人工资信息')
    os.chdir('个人工资信息')
else:
    os.chdir('个人工资信息')
window.mainloop()
