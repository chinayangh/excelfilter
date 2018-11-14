
<%

Dim xlsconn1,strs1ource,xlbook,xlsheet
Dim myconn1_Xsl,xlsrs1,sql,i
Set xlsconn1 = server.CreateObject("adodb.connection")
Set xlsrs1 = Server.CreateObject("Adodb.RecordSet")

myconn1_Xsl="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\inetpub\wwwroot\excel\sz\file\book2.xls;Extended Properties=Excel 8.0"
xlsconn1.open myconn1_Xsl

sql = "select '深圳' as 分校,学员X,学员名称,班级名称,课次,课程日期,上课时间,下课时间,教师名称,教学点,教室,母亲X,父亲X from [Sheet0$] where 班级名称 not like '短X班%' and 班级名称 not like '活动X%' and 班级名称 not like 'X试类%' and 班级名称 not like 'X断类%' and 班级名称 not like '%高中%' and 班级名称 not like '%双X%' and 班级名称 not like '%语文%' and 班级名称 not like '%大X堂%' and 教学点 not like '%雅仕X%' and 教学点 not like '%东X坊%' and 教学点 not like '%云X馆%' and 教学点 not like '%世X汇%'"

xlsconn1.Execute sql

xlsrs1.open sql,xlsconn1,1,1


if not xlsrs1.eof then
'Response.Write("<TABLE><TR>")
      For X = 0 To xlsrs1.Fields.Count - 1
         Response.Write("" & xlsrs1.Fields.Item(X).Name & ";")
      Next
      'Response.Write("</TR>")
Response.Write("<br>")
      xlsrs1.MoveFirst
end if



i=0
do While Not xlsrs1.EOF
	i=i+1
	a=xlsrs1("班级名称")
	b=xlsrs1("学员X")
	c=xlsrs1("学员名称")
	d=xlsrs1("课次")
	e=xlsrs1("课程日期")
	f=xlsrs1("上课时间")
	g=xlsrs1("下课时间")
	h=xlsrs1("教师名称")
	j=xlsrs1("教学点")
	k=xlsrs1("教室")
	m=xlsrs1("母亲X")
	n=xlsrs1("父亲X")
	p=xlsrs1("分校")
	response.write  p & ";" & b & ";" & c & ";" & a & ";" & d & ";" & e & ";" & f & ";" & g & ";" & h & ";" & j & ";" & k & ";" & m & ";" & n & "<br>" 
xlsrs1.MoveNext
loop


xlsconn1.close

%>
