
#获取路径下所有图片文件，并存入列表
import os
work_path = "图片\\"
pictures=[] # 存储文件夹内所有文件的路径（包括子目录内的文件）
for root, dirs, files in os.walk(work_path):
    path = [os.path.join(root, name) for name in files]
    pictures.extend(path)


from aip import AipOcr  #导入AipOcr模块，用于做文字识别
import time #时间模块
import requests #用于HTTP请求

APP_ID = '你申请的'
API_KEY = '你申请的'
SECRET_KEY = '你申请的'
client = AipOcr(APP_ID, API_KEY, SECRET_KEY)

#提交识别请求，并储存所有请求ID
for picture in pictures:
    pic = open(picture,'rb') #以二进制方式打开图片
    img = pic.read() #读取
    table = client.tableRecognitionAsync(img)    #调用表格识别模块
    request_id = table['result'][0]['request_id']
    
    #判断识别是否完成，直到完成才根据请求ID获取Excel下载路径
    result = client.getTableRecognitionResult(request_id)  #通过ID获取识别结果
    while result['result']['ret_msg'] != '已完成': #如果状态是“已完成”，才能获取下载地址
        time.sleep(2) #暂停2秒再刷新
        result = client.getTableRecognitionResult(request_id) #持续刷新，直到满足条件
        
    download_path = result['result']['result_data']
    
    #下载并将Excel文件名设为图片名
    excel_name = picture.split(".")[0] + ".xls" #让excel文件的名字与图片相同
    excel = requests.get(download_path) #抓取下载链接
    file = open(excel_name, 'wb') #新建excel文件
    file.write(excel.content) #写入excel文件并保存