import xlrd
import time 
def read_xls(filename):

    # 打开Excel文件
    data = xlrd.open_workbook(filename)

    data1=[]
    for i in range(4):
        table = data.sheets()[i]
    # 读取第一个工作表
    

    # 统计行数
        rows = table.nrows


       # 存放数据

        for v in range(1, rows):
            values = table.row_values(v)        
            data1.append(
                (
                    {
                    "time":str(values[0]),
                    "name":str(values[1]), # 这里我只需要字符型数据，加了str(),根据实际自己取舍
                    "item_type":str(values[2]),
                    "rank_type":str(values[3]).replace(".0", ""),
                    "id":str(values[4]),
                    "uigf_gacha_type":str(values[5]).replace(".0", ""),
                    }
                ) 
            )

    return data1

if __name__ == '__main__':
    uid=str(input("请输入您的UID："))
    now=time.strftime("%Y-%m-%d %H:%M:%S")
    xls=str(input("请将xls格式文件拖进来"))
    d1 = read_xls(xls)
    d2 = str(d1).replace("\'", "\"")    # 字典中的数据都是单引号，但是标准的json需要双引号
    
    info=str("\"info\":{\"uid\":"+uid+''',"lang":"zh-cn","export_time":"'''+now+''',"export_app":"Excel2JSON","export_app_version":"0.1.1","uigf_version":"v2.1"},''')
    d2 = "{"+info+"\"list\":" + d2 + "}"    # 前面的数据只是数组，加上外面的json格式大括号
    print(d2)
    # 可读可写，如果不存在则创建，如果有内容则覆盖
    jsFile = open("D:\\UIGF_"+uid+".json", "w+", encoding='utf-8')
    jsFile.write(d2)
    print("输出成功，文件在D盘根目录下")
    jsFile.close()
