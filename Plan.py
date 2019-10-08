#Plan类
import tkinter
from tkinter import messagebox
import time
from Plans.Plan.ExcelOperation import ExcelOperation
import ExcelData

#定义公共数据
data = None

#实例ExcelOperation
eo = ExcelOperation()

#创建一个GUI
win = tkinter.Tk()
win.title('Plans')
win.geometry('500x200+200+50')

#响应按键
def filePath():
    filePath=t.get()
    if filePath == '':
        messagebox.showinfo(title='title',message='文件路径为空！')
    return filePath

def btnRead():
    filepath =filePath()
    eo.nowExcel(filepath+'/Template')
    eo.writeExcel(filepath+'/Template','Template',ExcelData.templateData)


def lookData(data=data):
    messagebox.showinfo(title='读取的数据为：',message=data)

def btnWrite():
    filepath = filePath()
    l = eo.listchar()
    data = eo.excelFileRead(filepath+'Template')
    fd,d = eo.foumnData(data=data)
    eo.nowExcel(filepath+'/Tp'+str(time.localtime(time.time()).tm_year)+str(time.localtime(time.time()).tm_mon)+str(time.localtime(time.time()).tm_mday))
    eo.writeExcel(filepath+'/Tp'+str(time.localtime(time.time()).tm_year)+str(time.localtime(time.time()).tm_mon)+str(time.localtime(time.time()).tm_mday),'推移表',fd)
    eo.writeExcel(filepath+'/Tp'+str(time.localtime(time.time()).tm_year)+str(time.localtime(time.time()).tm_mon)+str(time.localtime(time.time()).tm_mday), 'BOM', d)



#控件
#ti = tkinter.Label(win,text='Plans',width=30,height=1)
l = tkinter.Label(win,text='工作路径：',width=20,height=1)
t = tkinter.Entry(show=None)
btnRead = tkinter.Button(win,text='生成模板',width=30,activebackground='#00ff00',command=btnRead)
btnWrite = tkinter.Button(win,text='写入公式',width=30,height=1,activebackground='#00ff00',command=btnWrite)
btnLookData = tkinter.Button(win,text='查看读取的数据',width=30,height=1,activebackground='#00ff00',command=lookData)

#显示
#ti.grid(row=0,column=1)
l.grid(row=2,column=0)
t.grid(row=2,column=1)
btnRead.grid(row=3,column=1)
btnWrite.grid(row=4,column=1)
btnLookData.grid(row=5,column=1)

#进入消息循环
win.mainloop()