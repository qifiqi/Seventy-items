import os

def Get_modify_time(file):
    return os.path.getmtime(file) #获取文件修改时间

path='文件'  #文件所在文件夹
files = [path+"\\"+i for i in os.listdir(path)] #获取文件夹下的文件名,并拼接完整路径
files.sort(key=Get_modify_time) #以文件修改时间为依据升序排序

seq = 1 #计数器，从1开始
for file in files:
    os.rename(file, os.path.join(path, str(seq) + ". "+ file.split("\\")[-1])) #重命名文件
    seq += 1