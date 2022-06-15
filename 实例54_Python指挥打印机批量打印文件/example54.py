#定义打印函数，以便重复调用
import win32api
def Print(fileName): 
    win32api.ShellExecute (
    0, #指定父窗口句柄，搞不懂
    "print", #指定操作，这里的"print"表示启动打印应用程序
    fileName, #要打印的文件名
    None, #打印机设置，若是"None"，则使用windows设置的默认打印机
    ".", #指定默认目录，照抄的，搞不懂
    0 #若fileName参数是一个可执行程序，则此参数指定程序窗口的初始显示方式，否则此参数应设置为0
    )

#获取待打印文件的路径
import os
path='文件'  #文件所在文件夹
files = [path+"\\"+i for i in os.listdir(path)] #获取文件夹下的文件名,并拼接完整路径

#批量打印
for file in files:
    Print(file)