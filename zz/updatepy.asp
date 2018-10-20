<%@LANGUAGE=Python%>
<html>
<head>
<script src="copy2clipboard.js?20180918" ></script>
</head>
<body>

<button onclick="select_all_and_copy(document.getElementById('tb'))">复制</button>



<%

Dim xlsconn1,strs1ource,xlbook,xlsheet
Dim myconn1_Xsl,xlsrs1,sql,i,sql2,xlsrs2
Set xlsconn1 = server.CreateObject("adodb.connection")
Set xlsrs1 = Server.CreateObject("Adodb.RecordSet")
Set xlsrs2 = Server.CreateObject("Adodb.RecordSet")

'myconn1_Xsl="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\inetpub\wwwroot\excel\gz\file\book2.xls;Extended Properties=Excel 8.0"
myconn1_Xsl="Provider=Microsoft.ACE.OLEDB.12.0;Data Source=c:\inetpub\wwwroot\excel\gz\file\book2.xls;Extended Properties=Excel 12.0"
xlsconn1.open myconn1_Xsl


sql = "select '广州' as 分校,学员uid,学员名称,班级名称,课次,课程日期,上课时间,下课时间,教师名称,count(教师名称) as num,教学点,教室,母亲电话,父亲电话 from [Sheet0$]   where 班级名称 not like '短期班%' and 班级名称 not like '活动类%' and 班级名称 not like '考试类%' and 班级名称 not like '诊断类%' and 班级名称 not like '%高中%' and 班级名称 not like '%双师%' and 班级名称 not like '%小学五年级语文%' and 班级名称 not like '%兴趣小组%' and 上课时间 like '%10:30%' or 上课时间 like '%08:00%' or 上课时间 like '%08:30%' or 上课时间 like '%12:00%' or 上课时间 like '%13:20%' or 上课时间 like '%14:50%' or 上课时间 like '%15:50%' or 上课时间 like '%18:00%' or 上课时间 like '%18:30%' group by 学员uid,学员名称,班级名称,课次,课程日期,上课时间,下课时间,教师名称,教学点,教室,母亲电话,父亲电话  order by 教师名称 "

sql2="select 教师名称,count(1) as num2 from [Sheet0$] group by 教师名称"

'xlsconn1.Execute sql

xlsrs1.open sql,xlsconn1,1,1

xlsrs2.open sql2,xlsconn1,1,1

if xlsrs1.eof then
	response.write "没有筛选到有效数据"

end if



if not xlsrs1.eof then
'Response.Write("<TABLE><TR>")
'      For X = 0 To xlsrs1.Fields.Count - 1
'         Response.Write("" & xlsrs1.Fields.Item(X).Name & ";")
'      Next
      'Response.Write("</TR>")
'Response.Write("<br>")
'      xlsrs1.MoveFirst

Response.Write("<TABLE id=tb ><TR>")
      For X = 0 To xlsrs1.Fields.Count - 1
         Response.Write("<TD >" & xlsrs1.Fields.Item(X).Name & "</TD>")
      Next
      Response.Write("</TR>")
      xlsrs1.MoveFirst

      While Not xlsrs1.EOF
         Response.Write("<TR>")
         For X = 0 To xlsrs1.Fields.Count - 1
            Response.write("<TD>" & xlsrs1.Fields.Item(X).Value)
         Next
         xlsrs1.MoveNext
         Response.Write("</TR>")
      Wend
      Response.Write("</TABLE>")
end if






'i=0
'do While Not xlsrs1.EOF
'	i=i+1
'	a=xlsrs1("班级名称")
'	b=xlsrs1("学员uid")
'	c=xlsrs1("学员名称")
'	d=xlsrs1("课次")
'	e=xlsrs1("课程日期")
'	f=xlsrs1("上课时间")
'	g=xlsrs1("下课时间")
'	h=xlsrs1("教师名称")
'	j=xlsrs1("教学点")
'	k=xlsrs1("教室")
'	m=xlsrs1("母亲电话")
'	n=xlsrs1("父亲电话")
'	p=xlsrs1("分校")
'	no=xlsrs1("num")
'	
'	response.write   p & ";" & b & ";" & c & ";" & a & ";" & d & ";" & e & ";" & f & ";" & g & ";" & h & ";" & no & ";" & j & ";" & k & ";" & m & ";" & n & "<br>" 
'	do while not xlsrs2.eof
'	no2=xlsrs2("num2")
'	response.write no2&"<br>"
'	xlsrs2.movenext
'	loop
'	xlsrs1.MoveNext
'	loop


xlsconn1.close

%>


</body>
</html>




