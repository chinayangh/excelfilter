
<%
Option Explicit
Response.Buffer = TRUE
Response.Expires = 0
Response.ContentType = "application/vnd.ms-excel"

Response.AddHeader "Content-Disposition", "attachment; filename = 0371data.xls"

%>
<html>
<head>
</head>
<body>
<%
Dim xlsconn1,xlbook,xlsheet
Dim myconn1_Xsl,xlsrs1,sql,i,sql2,xlsrs2
Set xlsconn1 = server.CreateObject("adodb.connection")
Set xlsrs1 = Server.CreateObject("Adodb.RecordSet")


'myconn1_Xsl="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\inetpub\wwwroot\excel\zz\file\book2.xls;Extended Properties=Excel 8.0"
myconn1_Xsl="Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\inetpub\wwwroot\excel\zz\file\book2.xls;Extended Properties=Excel 12.0"
xlsconn1.open myconn1_Xsl


sql = "select '郑州' as 分校,学员uid,学员名称,班级名称,课次,课程日期,上课时间,下课时间,教师名称,教学点,教室,母亲电话,父亲电话 from [Sheet0$]   where 班级名称 not like '短期班%' and 班级名称 not like '活动类%' and 班级名称 not like '考试类%' and 班级名称 not like '诊断类%' and 班级名称 not like '%高中%' and 班级名称 not like '%双师%'  and (上课时间 like '%10:30%' or 上课时间 like '%08:00%' or 上课时间 like '%08:30%' or 上课时间 like '%12:00%' or 上课时间 like '%13:20%' or 上课时间 like '%14:50%' or 上课时间 like '%15:50%' or 上课时间 like '%18:00%' or 上课时间 like '%18:30%')  order by 教师名称 "



'xlsconn1.Execute sql

xlsrs1.open sql,xlsconn1,1,1



if xlsrs1.eof then
	response.write "没有筛选到有效数据"

end if
%>
<table width="900" border="1">
<tr>
<th width="50"> <div align="center">分校 </div></th>
<th width="70"> <div align="center">学员uid </div></th>
<th width="70"> <div align="center">学员名称 </div></th>
<th width="200"> <div align="center">班级名称 </div></th>
<th width="50"> <div align="center">课次 </div></th>
<th width="100"> <div align="center">课程日期 </div></th>
<th width="70"> <div align="center">上课时间 </div></th>
<th width="70"> <div align="center">下课时间 </div></th>
<th width="70"> <div align="center">教师名称 </div></th>
<th width="70"> <div align="center">频次统计 </div></th>
<th width="70"> <div align="center">教学点 </div></th>
<th width="70"> <div align="center">教室 </div></th>
<th width="100"> <div align="center">母亲电话 </div></th>
<th width="100"> <div align="center">父亲电话 </div></th>
</tr>
<%
While Not xlsrs1.EOF
%>
<tr>
<td nowrap="nowrap"><div ><%=xlsrs1.Fields("分校").Value%></div></td>
<td nowrap="nowrap"><%=xlsrs1.Fields("学员uid").Value%></td>
<td nowrap="nowrap"><%=xlsrs1.Fields("学员名称").Value%></td>
<td nowrap="nowrap"><div ><%=xlsrs1.Fields("班级名称").Value%></div></td>
<td nowrap="nowrap"><%=xlsrs1.Fields("课次").Value%></td>
<td nowrap="nowrap"><%=xlsrs1.Fields("课程日期").Value%></td>
<td nowrap="nowrap"><%=xlsrs1.Fields("上课时间").Value%></td>
<td nowrap="nowrap"><%=xlsrs1.Fields("下课时间").Value%></td>
<td nowrap="nowrap"><%=xlsrs1.Fields("教师名称").Value%></td>
<td nowrap="nowrap"><%=xlsrs1.Fields("教学点").Value%></td>
<td nowrap="nowrap"><%=xlsrs1.Fields("教室").Value%></td>
<td nowrap="nowrap"><%=xlsrs1.Fields("母亲电话").Value%></td>
<td nowrap="nowrap"><%=xlsrs1.Fields("父亲电话").Value%></td>
</tr>
<%
xlsrs1.MoveNext
Wend
%>
</table>
    
</body>
</html>
<%
xlsrs1.Close()
xlsconn1.Close()
Set xlsrs1 = Nothing
Set xlsconn1 = Nothing
%>  