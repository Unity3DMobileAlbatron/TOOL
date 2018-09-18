# -*- coding: utf-8 -*-
#-*- coding:gbk-*-
import xlrd
import codecs
import os
homedir = os.getcwd()
homedir.replace("\\","/")
import time,datetime
path = homedir+"/青云志公测开服信息.xlsx"
endPath = homedir + "/"
allServerList = []
def open_excel(file):
    #try:
        data = xlrd.open_workbook(filename=file,encoding_override="utf-8")
        return data
    #except Exception,e:
        #print str(e)

def write_table(file,colnames,myList):
  s_list = []
  for rownum in range(0,len(myList)):
    row = myList[rownum]

    if row[7].upper()!="TRUE":
      if row[8] !="" and row[8] !="0" and str(row[8]).strip():
        llist = row[8].split(",")
        print(llist)
        for x in range(0,len(llist)):
          serverid = llist[x]
          for i in range(0,len(allServerList)):
            tmp = allServerList[i];
            if tmp[7].upper()=="TRUE" and  serverid!="" and tmp[5]!="" and (int(tmp[5]) == int(serverid)):
              if tmp[3] != row[3]:
                print(str(tmp[1])+" ip=" + str(tmp[3]))
                print(str(row[1])+" ip=" + str(row[3]))
                print("error 内网ip不一致")
                os.system('pause')
                return
              if tmp[4] != row[4]:
                print(str(tmp[1])+" ip=" + str(tmp[4]))
                print(str(row[1])+" ip=" + str(row[4]))
                print("error 外网ip不一致")
                os.system('pause')
                return
              row[1] = row[1] + ";" + tmp[1]
      s_list.append(row)

  for rownum in range(0,len(s_list)):
   ip = "a"
   row = s_list[rownum]
   for i in range(1,len(colnames)):
    if colnames[i] == "开服时间":
      tmp = str(row[i]);
      if tmp != "0" and tmp != "":
           for i in range(len(colnames)):
              if colnames[i] == "":
                continue;
                pass
              #name
              if colnames[i] == "游戏组名":
                tmp = str(row[i]);
                file.write(tmp);
                file.write("|");
              pass
              if colnames[i] == "机器名":
                tmp = str(row[i]);
                file.write(tmp);
                file.write("|");
              pass
              if colnames[i] == "内网ip":
                tmp = str(row[i]);
                file.write(tmp);
                file.write("|");
              pass
              if colnames[i] == "外网ip":
                tmp = str(row[i]);
                file.write(tmp);
                file.write("|");
              pass

              if colnames[i] == "serverid":
                if row[i] != "0" and row[i] != 0  and row[i] != "":
                  tmp = str(int(row[i]));
                  file.write(tmp);
                file.write("|");
              pass

              if colnames[i] == "开始端口":
                if row[i] != "0" and row[i] != 0  and row[i] != "":
                  tmp = str(int(row[i]));
                  file.write(tmp);
                file.write("|");
              pass

              if colnames[i] == "合服":
                tmp = str(row[i]);
                if tmp == "" or tmp == "0":
                  file.write("0");
                else:
                  file.write(tmp.upper());
                file.write("|");
              pass
              if colnames[i] == "合服列表":
                tmp = str(row[i]);
                file.write(tmp);
                file.write("\n");
              pass
  print("-----------------------------------")

def insertServerList(SheetName):
  data = open_excel(path)
  table = data.sheet_by_name(SheetName)
  nrows = table.nrows #行数 
  colnames = table.row_values(0);
  for rownum in range(1,nrows):
     ip = "a"
     row = table.row_values(rownum)
     allServerList.append(row)
#@param fileName:输出文件名称
#@param SheetName:excel表名

def excel_table_byname(fileName,SheetName ,index,openStatus="w"):
    data = open_excel(path)
    table = data.sheet_by_name(SheetName)
    nrows = table.nrows #行数 
    colnames = table.row_values(0);

    file = codecs.open(endPath+fileName+".txt",openStatus,"utf-8")

    myList = [];
    for rownum in range(index,nrows):
       ip = "a"
       row = table.row_values(rownum)
       myList.append(row)
    write_table(file,colnames,myList)
    file.close();       
def main():
  #android
  insertServerList(SheetName="混服")
  insertServerList(SheetName="应用宝")
  insertServerList(SheetName="ios+安卓官方")

  excel_table_byname(SheetName="混服",fileName="jinshanyun",index=1,openStatus = "w");
  excel_table_byname(SheetName="应用宝",fileName="tenxunyun",index=1,openStatus = "w");
  excel_table_byname(SheetName="ios+安卓官方",fileName="aliyun",index=1,openStatus = "w");
if __name__=="__main__":

    main()



