<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>

<%
'If IsEmpty(Session("username")) Then
'	response.write "<script>window.location.href='index.html'</script>"
'else
set conn1=server.createobject("adodb.connection")
conn1str="provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\inetpub\wwwroot\kdinfo.mdb"
conn1.Open conn1str
Set rs1=server.CreateObject("adodb.recordset")


Dim xlsconn1,strs1ource,xlbook,xlsheet,i
Dim myconn1_Xsl,xlsrs1,sql,sql2
Set xlsconn1 = server.CreateObject("adodb.connection")
Set xlsrs1 = Server.CreateObject("Adodb.RecordSet")


i=0
myconn1_Xsl="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\inetpub\wwwroot\book2.xls;Extended Properties=Excel 8.0"
xlsconn1.open myconn1_Xsl

sql = "Select * from [Sheet1$] where 运单号 is not null"

xlsrs1.open sql,xlsconn1,1,1


a=xlsrs1("运单号")

Function checkStr(Chkstr) 
dim Str:Str=Chkstr 
if isnull(Str) then 
   checkStr = "" 
   exit Function 
else 
   Str=replace(Str,"’","") 
   Str=replace(Str,";","") 
   Str=replace(Str,"--","") 
   checkStr=Str 
end if 
End Function

If xlsrs1.eof Then

 elseif  not conn1.execute("select * from kdinfo where kd_number like '"&a&"'").eof Then
 danhao=conn1.execute("select * from kdinfo where kd_number like '"&a&"'").Fields("kd_number")

  	Response.write danhao&"快递单号已经存在相同的记录<br>"

  elseif not xlsrs1.eof Then

	do While Not xlsrs1.EOF
		i=i+1
		

	a=trim(xlsrs1("运单号"))


	b=xlsrs1("快递公司")

	c=xlsrs1(9)


	d=xlsrs1("暑假课程简名")


	e=xlsrs1("地址")


	f=trim(xlsrs1(5))


	g=trim(xlsrs1("收件人"))

	if trim(xlsrs1("学员编码"))="" then
	h=dbnull.value
	else
	h=trim(xlsrs1("学员编码"))
	end if 
	

	k=xlsrs1("班级分校")


	j=xlsrs1("邮寄日期")

		
		Response.write i&":"&a&"<br>"
		For X = 0 To xlsrs1.Fields.Count - 1
				
			Next
			
           sql2="insert into kdinfo(kd_number,company,sender,lesson_name,address,phone,receivername,stuNumber,school,kd_date) values('"&a&"','"&b&"','"&c&"','"&d&"','"&e&"','"&f&"','"&g&"','"&h&"','"&k&"','"&j&"')"
  			conn1.execute(sql2)
  			
  			
  		xlsrs1.MoveNext
        


     loop
       

   End If
xlsrs1.close
conn1.close

Response.CharSet = "GB2312"
Response.write "共导入<font color='red'>" & i & "</font>条记录.<br>" 

set xlsconn1=nothing
'end if
%>
<!--<script>alert("单号导入成功！");</script>-->
