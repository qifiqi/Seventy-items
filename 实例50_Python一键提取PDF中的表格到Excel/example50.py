import os
import pdfplumber
from openpyxl import Workbook    

path='PDF'  #文件所在文件夹
files = [path+"\\"+i for i in os.listdir(path)] #获取文件夹下的文件名,并拼接完整路径
key_words = "主要会计数据"

for file in files:
    with pdfplumber.open(file) as p:
        wb = Workbook() #新建excel工作簿
        wb.remove(wb.worksheets[0])#删除工作簿自带的工作表
        
        #获取关键词所在页及下一页的页码
        pages_wanted = []
        for index,page in enumerate(p.pages): #从0开始给所有页编号
            if key_words in page.extract_text():
                pages_wanted.append(index)
                pages_wanted.append(index+1)
                break
        
        #提取指定页码里的表格
        for i in pages_wanted:     
            page = p.pages[i]
            tables = page.extract_tables() #读取表格
            if tables: #判断是否存在表格，若不存在，则不执行下面的语句
                ws = wb.create_sheet(f"Sheet{i+1}") #新建工作表，表名的编号与表在PDF中的页码一致
                for table in tables: #遍历所有列表
                    for row in table: #遍历列表中的所有子列表，里面保存着行数据
                        ws.append(row) #写入excel表
        wb.save("Excel\\{}.xlsx".format(file.split("\\")[1].split(".")[0]))