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

def open_excel(file):
    #try:
        data = xlrd.open_workbook(filename=file,encoding_override="utf-8")
        return data
    #except Exception,e:
        #print str(e)

#@param fileName:输出文件名称
#@param SheetName:excel表名

def excel_table_byname(fileName,SheetName ,serverListName,index,nameindx,openStatus="w"):
    data = open_excel(path)
    table = data.sheet_by_name(SheetName)
    nrows = table.nrows #行数 
    colnames = table.row_values(0);
    text_ = ["游戏组名","外网ip","开始端口"];
    #name = ["name","test","isNew","addresses","host","port"];
    file = codecs.open(endPath+fileName+".txt",openStatus,"utf-8")
    find = False
    for rownum in range(index,nrows):
         ip = "a"
         row = table.row_values(rownum)
         if row[0]!= "":
            if find:
                file.write("}\n");
                nameindx = nameindx + 1
                find = False
            find = True
            file.write("local ");
            file.write(serverListName[nameindx])
            file.write(" = ");
            file.write("{\n");
         for i in range(len(colnames)):
          if colnames[i] == "开服时间":
            tmp = str(row[i]);
            if tmp != "0" and tmp != "":
                #file.write("");
                 
                 file.write("{");
                 for i in range(len(colnames)):
                    if colnames[i] == "":
                      continue;
                      pass
                    #name
                    if colnames[i] == "游戏组名":
                      file.write("name");
                      file.write(" = ");
                      file.write("\"");
                      tmp = str(row[i]);
                      if tmp == "0" or tmp == "0.0":
                        file.write("");
                      else:
                        file.write(tmp);
                      file.write("\",");
                    pass

                    if colnames[i] == "外网ip":
                        tmp = str(row[i]);
                        if tmp != "0" and tmp != "0.0" and ip!="":
                          ip = tmp
                    if colnames[i] == "开始端口" and ip!="a":
                      file.write("addresses");
                      file.write(" = ");
                      file.write("{");

                      
                      for index in range(0,5):
                        file.write("{");
                        file.write("host");
                        file.write(" = ");
                        file.write("\"");
                        file.write(ip);
                        file.write("\",");
                        file.write("port");
                        file.write(" = ");
                        if row[i]!="" and str(row[i]).strip():
                          file.write(str(int(row[i])+index));
                        else:
                          file.write("0");
                        file.write("},");
                        
                      file.write("{");
                      file.write("host");
                      file.write(" = ");
                      file.write("\"");
                      file.write(ip);
                      file.write("\",");
                      file.write("port");
                      file.write(" = ");
                      file.write(str(int(row[i])+5));
                      file.write("}");
                      file.write("},");
                    pass
                    if colnames[i] == "是否新服":
                      file.write("isNew");
                      file.write(" = ");
                      if row[i]=="是":
                        file.write("true");
                      else:
                        file.write("false");
                    pass

                 file.write("},\n");

    file.write("}\n");
    file.close();       
def main():
  #android
  serverListName = ["serverlist","yhlm_serverlist","dhf_serverlist","yingyongbao_serverlist"]
  excel_table_byname(SheetName="混服",fileName="android_server",index=1,nameindx=0,serverListName = serverListName,openStatus = "w");
  excel_table_byname(SheetName="应用宝",fileName="android_server",index=1,nameindx=3,serverListName = serverListName,openStatus = "a+");
  file = codecs.open(endPath+"android_server"+".txt","a+","utf-8")
  file.write("local logserver = {host=\"10.12.3.122\", port=10031}\n");
  file.write("return\n");
  file.write("{\n");
  for serverIndex in range(0,len(serverListName)):
    file.write("    ");
    file.write(serverListName[serverIndex]);
    file.write(" = ");
    file.write(serverListName[serverIndex]);
    file.write(",\n");
    pass
  file.write("    ");
  file.write("logserver = logserver");
  file.write("\n");
  file.write("}");
  file.close();
  #ios
  ios_serverlistName = ["ios_serverlist"] #ios 暂时用大混服的索引
  excel_table_byname(SheetName="ios+安卓官方",fileName="ios_server",index=1,nameindx=0,serverListName = ios_serverlistName,openStatus = "w");
  file = codecs.open(endPath+"ios_server"+".txt","a+","utf-8")
  file.write("local recommendserver = 2\n");
  file.write("local logserver = {host=\"10.12.3.122\", port=10031}\n");
  file.write("return\n");
  file.write("{\n");
  '''
  for serverIndex in range(0,len(ios_serverlistName)):
    file.write("    ");
    file.write(ios_serverlistName[serverIndex]);
    file.write(" = ");
    file.write(ios_serverlistName[serverIndex]);
    file.write(",\n");
    pass
  '''
  file.write("    ");
  file.write("dhf_serverlist = ios_serverlist");
  file.write(",\n");
  file.write("    ");
  file.write("logserver = logserver");
  file.write("\n");
  file.write("}");
  file.close();


if __name__=="__main__":

    main()



