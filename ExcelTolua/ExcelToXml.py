# -*- coding: utf-8 -*-
#-*- coding:gbk-*-
import xlrd
import codecs
import os
import types 
homedir = os.getcwd()
homedir.replace("\\","/")
import time,datetime
path = homedir+"/青云志公测开服信息.xlsx"
endPath = homedir + "/"

def open_excel(file):
    #try:
        data = xlrd.open_workbook(filename=file,encoding_override="utf-8")
        return data
    #except Exception,e:
        #print str(e)
#@param fileName:输出文件名称
#@param SheetName:excel表名

def excel_table_byname(fileName,SheetName,cloud,content ,content_py ,index,nameindx,openStatus="w"):
    data = open_excel(path)
    table = data.sheet_by_name(SheetName)
    nrows = table.nrows #行数 
    colnames = table.row_values(0);
    #name = ["name","test","isNew","addresses","host","port"];
    elementname = ["0","group", "machine","intranetIp","extranetIp","serverid","begainPort","merge","mergelist","serverCloudName","openTime","auanyIp","auanyPort","servicedIp","servicedPort","servicedglobalid","configure" , "newServer"]
    #               游戏组名 机器名     内网IP       外网ip      serverid   开始端口          和服    和服列表     服务器厂商       开服时间   auanyIp   auany端口   servicedIp   serviced端口   servicedglobalid    服务器配置   是否新服
    # openType 是否开服
    file = codecs.open(endPath+fileName+".xml","a+","utf-8")
    find = False
    file.write("    ")
    file.write("<");
    file.write("cloudType");
    file.write(" ")
    file.write("name=");
    file.write("\"")
    file.write(cloud[1]);
    file.write("\"")
    file.write(" ")
    file.write("comment=")
    file.write("\"")
    file.write(cloud[0]);
    file.write("\">\n")
    
    for rownum in range(index,nrows):
         ip = "a"
         row = table.row_values(rownum)
         if row[0]!= "":
            if find:
                file.write("      ")
                file.write("</serverType>\n");
                nameindx = nameindx + 1
                find = False
            find = True
            file.write("      ")
            file.write("<");
            file.write("serverType");
            file.write(" ")
            file.write("name=");
            file.write("\"")
            file.write(content_py[nameindx]);
            file.write("\"")
            file.write(" ")
            file.write("comment=")
            file.write("\"")
            file.write(content[nameindx]);
            file.write("\">\n")
         file.write("         ")
         file.write("<");
         file.write("serverlist");
         file.write(" ")
         for i in range(len(colnames)-1):
            
            tmp = row[i+1]
            if tmp!="" and str(tmp).strip():
                tmp = row[i+1];
            else:
                tmp = "0";
            if colnames[i+1] == "开服时间":
                if tmp != "0" and tmp != "":
                  now = xlrd.xldate.xldate_as_datetime(row[i+1], 0)
                  file.write(elementname[i+1])
                  file.write("=")
                  file.write("\"")
                  file.write(str(now).strip())
                  file.write("\"")
                  file.write(" ")
                  file.write("openType") #是否开服
                  file.write("=")
                  file.write("\"")
                  file.write("1")
                  file.write("\"")
                  file.write(" ")
                else: 
                  file.write(elementname[i+1])
                  file.write("=")
                  file.write("\"")
                  file.write("0")
                  file.write("\"")
                  file.write(" ")
                  file.write("openType") #是否开服
                  file.write("=")
                  file.write("\"")
                  file.write("0")
                  file.write("\"")
                  file.write(" ")
            elif colnames[i+1] == "服务器配置":
                if tmp == "高配":
                  file.write(elementname[i+1])
                  file.write("=")
                  file.write("\"")
                  file.write("1")
                  file.write("\"")
                  file.write(" ")
                else:
                  file.write(elementname[i+1])
                  file.write("=")
                  file.write("\"")
                  file.write("0")
                  file.write("\"")
                  file.write(" ")
            elif colnames[i+1] == "是否新服":
                if tmp == "是":
                  file.write(elementname[i+1])
                  file.write("=")
                  file.write("\"")
                  file.write("1")
                  file.write("\"")
                  file.write(" ")
                else:
                  file.write(elementname[i+1])
                  file.write("=")
                  file.write("\"")
                  file.write("0")
                  file.write("\"")
                  file.write(" ")
            elif colnames[i+1] == "合服":
                print(tmp)
                if tmp == "" or tmp == "0":
                  file.write(elementname[i+1])
                  file.write("=")
                  file.write("\"")
                  file.write("0")
                  file.write("\"")
                  file.write(" ")
                else:
                  file.write(elementname[i+1])
                  file.write("=")
                  file.write("\"")
                  file.write(tmp.upper())
                  file.write("\"")
                  file.write(" ")
            else:
              if type(tmp) is float:
                file.write(elementname[i+1])
                file.write("=")
                file.write("\"")
                file.write(str(int(tmp)))
                file.write("\"")
                file.write(" ")
              else:
                file.write(elementname[i+1])
                file.write("=")
                file.write("\"")
                file.write(str(tmp).strip())
                file.write("\"")
                file.write(" ")
         file.write("         ")
         file.write("/>\n");
    file.write("      ")
    file.write("</serverType>\n");
    file.write("    ")
    file.write("</cloudType>\n");
    file.close();       
def main():
  #android
  serverListName = ["serverlist","yhlm_serverlist","dhf_serverlist","yingyongbao_serverlist"]
  #excel_table_byname(SheetName="混服",fileName="android_server",index=1,nameindx=0,serverListName = serverListName,openStatus = "w");
  #excel_table_byname(SheetName="应用宝",fileName="android_server",index=1,nameindx=3,serverListName = serverListName,openStatus = "a+");
  name = "serverlist"
  file = codecs.open(endPath+name+".xml","w","utf-8")
  file.write("<?xml version = \"1.0\" encoding = \"UTF-8\" ?>\n");
  file.write("  ");
  file.write("<root>\n");
  file.close();
  content = ["官方","硬盒联盟","大混服"]
  content_py = ["gf","yhlm","dhf"]
  cloud = ["金山云","jsy"]
  excel_table_byname(SheetName="混服",fileName=name,cloud =cloud,content = content,content_py=content_py,index=1,nameindx=0,openStatus = "a+");
  content = ["腾讯服"]
  content_py = ["txf"]
  cloud = ["腾讯云","txy"]
  excel_table_byname(SheetName="应用宝",fileName=name,cloud =cloud,content = content,content_py=content_py,index=1,nameindx=0,openStatus = "a+");
  content = ["ios服"]
  content_py = ["ios"]
  cloud = ["阿里云","aly"]
  excel_table_byname(SheetName="ios+安卓官方",fileName=name,cloud =cloud,content = content,content_py=content_py,index=1,nameindx=0,openStatus = "a+");
  file = codecs.open(endPath+name+".xml","a+","utf-8")
  file.write("  ");
  file.write("</root>\n");
  


if __name__=="__main__":
    main()



