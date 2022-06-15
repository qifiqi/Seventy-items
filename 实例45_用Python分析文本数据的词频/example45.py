import os
path='主要业务'  #文件所在文件夹
files = [path+"\\"+i for i in os.listdir(path)] #获取文件夹下的文件名,并拼接完整路径

import jieba
import pandas as pd

for file in files:

    txt = open(file, "r", encoding="utf-8").read()
    words = jieba.lcut(txt)
    wordsDict = {} #新建字典用于储存词及词频
    for word in words:
        if len(word) == 1: #单个的字符不作为词放入字典
            continue
        else:
            wordsDict.setdefault(word, 0) #设置词的初始出现次数为0
            wordsDict[word] +=1 #对于重复出现的词，每出现一次，次数增加1

    stopWords = ["2019","不断","持续","主要","企业","产品","业务","公司","行业","000","用于","情况","方面","一种","要求","对于","进行","一般","212","实现","处理","通过","投入","随着"]
    for word in stopWords:
        if word in wordsDict:
            del wordsDict[word]

    wordsDict_seq = sorted(wordsDict.items(),key=lambda x:x[1], reverse=True) #按字典的值降序排序

    df = pd.DataFrame(wordsDict_seq,columns=['词','次数'])
    df.to_excel("词频//{}.xlsx".format(file.split("\\")[1][:-4]),index = False) #存为Excel时去掉index索引列