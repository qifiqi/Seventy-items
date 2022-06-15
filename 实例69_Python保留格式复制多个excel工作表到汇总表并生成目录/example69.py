import os #用于获取目标文件所在路径
import win32com

path=os.getcwd()+"\\文件\\" # 文件夹绝对路径
files=[]
for file in os.listdir(path):
    if file.endswith(".xls") or file.endswith(".xlsx"): #只获取".xls"后缀的文件
        files.append(path+file) 


excel_app = win32com.client.Dispatch("Excel.Application")
excel_app.Visible = False  # 不显示Excel文件
excel_app.DisplayAlerts = False

#新建excel工作簿
wb = excel_app.Workbooks.Add()
wb.SaveAs(os.getcwd() + "\\汇总.xlsx")
ws = wb.Worksheets(1)
ws.Name = "目录"
for i in range(len(files)):
    file_name = files[i].split("\\")[-1].split(".")[0]
    ws.Range("A"+str(i+1)).Value = file_name
    #读取子文件
    wb_sub = excel_app.Workbooks.Open(files[i])
    ws_sub = wb_sub.ActiveSheet # #获取活动工作表

    ws_sub.Copy(ws) #复制工作表到汇总表
    wb.ActiveSheet.Name = file_name #更改工作表名
    wb_sub.Close()
    print(f"已复制文件 {file_name}")


#将“目录”工作表移动到最前面
first_sheet = files[0].split("\\")[-1].split(".")[0]
wb.Worksheets("目录").Move(wb.Worksheets(first_sheet))    

wb.Save()
wb.Close()
excel_app.Quit()