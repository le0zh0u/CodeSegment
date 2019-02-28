import xdrlib ,sys
import xlrd
excel_file_name = 'a.xlsx'
col_index = [0]
sheet_name = ''
reault_file = 'employee.sql'
key_column = 'para1'
value_column = 'para2'
mysql_table_name = 'table_name'
def open_excel(file= 'a.xlsx'):
    try:
        data = xlrd.open_workbook(file)#打开excel文件
        return data
    except Exception,e:
        print str(e)

def excel_table_bycol(file='a.xlsx',colindex=[0],table_name='Sheet1'):
    data = open_excel(file)
    table = data.sheet_by_name(table_name)#获取excel里面的某一页
    nrows = table.nrows#获取行数
    colnames = table.row_values(0)#获取第一行的值，作为key来使用，对于不同的excel文件可以进行调整
    list = []
    #（1，nrows）表示取第一行以后的行，因为第一行往往是表头
    for rownum in range(1,nrows):
         row = table.row_values(rownum)
         if row:
              app = {}
              for i in colindex:
                   app[str(colnames[i]).encode("utf-8")] = str(row[i]).encode("utf-8")#将数据填入一个字典中，同时对数据进行utf-8转码，因为有些数据是unicode编码的
              list.append(app)#将字典加入列表中去
    return list
def main():
   #colindex是一个数组，用来选择读取哪一列，因为往往excel中的一小部分才是我们需要的
   tables = excel_table_bycol(fiel=excel_file_name, colindex=col_index,table_name=sheet_name)
   file = open(reault_file,'w')#创建sql文件，并开启写模式
   for row in tables:
        if row[key_column] == '' || row[value_column] == '':
            print "数据有误"
        else 
            print "update %s set para1='%s' where para2='%s';\n"%(mysql_table_name, row[key_column],row[value_column])
            # file.write("update table_name set para1='%s' where para2='%s';\n"%(row['para1'],row['para2']))#往文件里写入sql语句
if __name__=="__main__":
    main()