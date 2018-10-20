#!/usr/bin/python
print ('Status: 200 OK')
print ('Content-type: text/html')
print ()

import cgi
import cgitb
cgitb.enable()

#import win32com.client as win32
#fname = "C:\\inetpub\\wwwroot\\order.xls"
#excel = win32.gencache.EnsureDispatch('Excel.Application')
#wb = excel.Workbooks.Open(fname)

#wb.SaveAs(fname+"x", FileFormat = 51)    #FileFormat = 51 is for .xlsx extension
#wb.Close()                               #FileFormat = 56 is for .xls extension
#excel.Application.Quit()


#import pyexcel as p

#p.save_book_as(file_name='C:\\inetpub\\wwwroot\\order.xls',
#               dest_file_name='C:\\inetpub\\wwwroot\\order.xlsx')


import xlwt
import codecs

def Txt_to_Excel(inputTxt,sheetName,start_row,start_col,outputExcel):
  fr = codecs.open(inputTxt,'r')
  wb = xlwt.Workbook(encoding = 'utf-8')
  ws = wb.add_sheet(sheetName)

  line_number = 0#记录有多少行，相当于写入excel时的i，
  row_excel = start_row
  try:
          jiaoshi=[]
          for line in fr :
              line_number +=1
              row_excel +=1
              line = line.strip()
              line = line.split('\t')
              #print(line,'<br>')
              #print(line[8])
              #jiaoshi.append(line[8].split(" "))
              
              len_line = len(line)#list中每一行有多少个数，相当于写入excel中的j
              col_excel = start_col
              for j in range(len_line):
                        
                        #print(j)
                        #print (line[j])
                        ws.write(row_excel,col_excel,line[j])
                        col_excel +=1
                        wb.save(outputExcel)

  except Exception as e:
      print('出错啦！')
      print(str(e))

  print('<a href=file/020data_conver.xls>点此下载转换文件</a>')
  print('&nbsp&nbsp<a href=javascript:history.back(-1)>返回上一页</a>')
  

if __name__=='__main__':
 sheetName = 'Sheet2'#需要写入excel中的Sheet2中，可以自己设定
 start_row = -1 #从第7行开始写
 start_col = 0 #从第3列开始写
 inputfile = 'C:\\inetpub\\wwwroot\\excel\\gz\\file\\020data.xls' #输入文件
 outputExcel = 'C:\\inetpub\\wwwroot\\excel\\gz\\file\\020data_conver.xls' #输出excel文件
 Txt_to_Excel(inputfile,sheetName,start_row,start_col,outputExcel)
