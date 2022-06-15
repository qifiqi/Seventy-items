import os
from datetime import datetime
from time import time, localtime ,strftime
from tkinter import Tk,Entry,Button,Listbox,X,Y,END,Scrollbar,RIGHT,BOTTOM,HORIZONTAL,StringVar,Label
from tkinter.filedialog import askdirectory

class MainGUI():
    def __init__(self):
        myWindow = Tk()
        myWindow.title("批量新建文件夹")
        #设置窗口大小
        myWindow.geometry('590x400')
        #增加标签
        self.label_1 = Label(myWindow, text=' 目标目录:')
        self.label_1.place(x=10, y=10,width=70, height=30)
        self.label_2 = Label(myWindow, text='文件夹数量:')
        self.label_2.place(x=10, y=50,width=70, height=30)

        #增加文本框
        addr = StringVar(value='C:\\Users\\xxxx\\Desktop') #文本框默认显示的内容
        self.input_entry = Entry(myWindow, highlightcolor='red', highlightthickness=1, textvariable=addr)
        self.input_entry.place(x=80, y=10,width=410, height=30)
        self.btn_in = Button(myWindow, text='选择目录',command = self.select_dir1, width=10, height=1) 
        self.btn_in.place(x=500,y=10)
        
        folder_quantity = str(self.get_folder_qty()) #从日志文件`log.txt`中读取最近使用过的文件夹数量
        def_qty = StringVar(value = folder_quantity)
        self.folderQty_entry = Entry(myWindow, highlightcolor='blue', highlightthickness=1, textvariable=def_qty)
        self.folderQty_entry.place(x=80, y=50,width=410, height=30)
        self.btn_exe = Button(myWindow, text='执行新建',command = self.create_folders, width=10, height=1)
        self.btn_exe.place(x=500,y=50)
        
        #增加列表框
        self.result_show = Listbox(myWindow,bg='Azure') 
        self.result_show.place(x=10,y=100, width=570, height=290)
        self.sbY = Scrollbar(self.result_show,command=self.result_show.yview)#在列表框中增加Y轴滚动条
        self.sbY.pack(side=RIGHT,fill=Y)
        self.result_show.config(yscrollcommand = self.sbY.set)
        self.sbX = Scrollbar(self.result_show,command=self.result_show.xview,orient = HORIZONTAL)#在列表框中增加X轴滚动条
        self.sbX.pack(side=BOTTOM,fill=X)
        self.result_show.config(xscrollcommand = self.sbX.set)
        
        myWindow.mainloop()
        
    def select_dir1(self):
        self.input_entry.delete(0, END)
        self.input_entry.insert(0, askdirectory(initialdir= "D:\\"))
        
    def create_folders(self):
        date = self.get_current_date() #获取日期
        qty = int(self.folderQty_entry.get()) #获取文本框中文件夹数量，并转为整数
        for i in range(1,qty+1):
            folder = self.input_entry.get() + "\\" + date + '-' + str(i)
            # 判断是否已经存在该目录
            if not os.path.exists(folder):
                # 目录不存在，进行创建操作
                os.makedirs(folder) #使用os.makedirs()方法创建目录
                f = f"目录新建成功：{folder}" # 创建一个显示项
                self.result_show.insert("end", f) #将结果添加到列表框中
            else:
                f = f"目录已存在：{folder}" # 创建一个显示项
                self.result_show.insert("end", f) #将结果添加到列表框中
        f = "-"*100 #创建分割线
        self.result_show.insert("end", f) # 将分割线添加到列表框
        f = f"程序运行完成，请关闭窗口退出."# 创建一个显示项
        self.result_show.insert("end", f) # 将结果添加到列表框
        f = " "*100
        self.result_show.insert("end", f) # 将以上空格添加到列表框
        self.save_recent_folder_qty() #保存最新的文件夹数量
        
    def get_current_date(self):
        time_stamp = time()  
        local_time = localtime(time_stamp)  
        str_time_month = int(strftime('%m', local_time))
        str_time_day = int(strftime('%d', local_time))
        return str(str_time_month)+"."+str(str_time_day)
    
    def get_folder_qty(self):
        '''从log.txt文件中获取最近的文件夹数量，若没有则返回0'''
        log_file = os.getcwd() + "\\log.txt"
        if os.path.exists(log_file):
            with open(log_file, "r") as f:
                qty = f.readline()
                return int(qty)
        else:
            return 0
        
    def save_recent_folder_qty(self):
        '''保存最近的文件夹数量'''
        log_file = os.getcwd() + "\\log.txt"
        with open(log_file, "w") as f:
            recent_qty = str(self.folderQty_entry.get())
            f.write(recent_qty) 
            
if __name__ == "__main__":
    MainGUI()