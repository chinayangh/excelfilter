<%@ Language=VBScript %>
<%
Option Explicit
Response.Buffer = TRUE
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename = 0755data.xls"
%>
<html>
<head>
</head>
<body>
<%
Dim xlsconn1,strs1ource,xlbook,xlsheet
Dim myconn1_Xsl,xlsrs1,sql,i,sql2,xlsrs2
Set xlsconn1 = server.CreateObject("adodb.connection")
Set xlsrs1 = Server.CreateObject("Adodb.RecordSet")
Set xlsrs2 = Server.CreateObject("Adodb.RecordSet")

'myconn1_Xsl="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\inetpub\excel\sz\file\book2.xls;Extended Properties=Excel 8.0"
myconn1_Xsl="Provider=Microsoft.ACE.OLEDB.12.0;Data Source=c:\inetpub\excel\sz\file\book2.xls;Extended Properties=Excel 12.0"
xlsconn1.open myconn1_Xsl

sql = "select '����' as ��У,ѧԱuid,ѧԱ����,�༶����,�δ�,�γ�����,�Ͽ�ʱ��,�¿�ʱ��,��ʦ����,��ѧ��,����,ĸ�׵绰,���׵绰 from [Sheet0$] where �༶���� not like '���ڰ�%' and �༶���� not like '���%' and �༶���� not like '������%' and �༶���� not like '�����%' and �༶���� not like '%����%' and �༶���� not like '%˫ʦ%' and �༶���� not like '%����%' and �༶���� not like '%����%' and ��ѧ�� not like '%������%' and ��ѧ�� not like '%������%' and ��ѧ�� not like '%���й���%' and ��ѧ�� not like '%���ͻ�%' and �༶���� not like '%�˰�ȡ��%' order by ��ʦ����"

sql2="select ��ʦ����,count(1) as num2 from [Sheet0$] group by ��ʦ����"

'xlsconn1.Execute sql

xlsrs1.open sql,xlsconn1,1,1

xlsrs2.open sql2,xlsconn1,1,1

if xlsrs1.eof then
	response.write "û��ɸѡ����Ч����"

end if
%>
<table width="1000" border="1">
<tr>
<th width="50"> <div align="center">��У </div></th>
<th width="70"> <div align="center">ѧԱuid </div></th>
<th width="70"> <div align="center">ѧԱ���� </div></th>
<th width="200"> <div align="center">�༶���� </div></th>
<th width="50"> <div align="center">�δ� </div></th>
<th width="100"> <div align="center">�γ����� </div></th>
<th width="70"> <div align="center">�Ͽ�ʱ�� </div></th>
<th width="70"> <div align="center">�¿�ʱ�� </div></th>
<th width="70"> <div align="center">��ʦ���� </div></th>
<th width="70"> <div align="center">��ѧ�� </div></th>
<th width="70"> <div align="center">���� </div></th>
<th width="100"> <div align="center">ĸ�׵绰 </div></th>
<th width="100"> <div align="center">���׵绰 </div></th>
</tr>
<%
While Not xlsrs1.EOF
%>
<tr>
<td nowrap="nowrap"><div ><%=xlsrs1.Fields("��У").Value%></div></td>
<td nowrap="nowrap"><%=xlsrs1.Fields("ѧԱuid").Value%></td>
<td nowrap="nowrap"><%=xlsrs1.Fields("ѧԱ����").Value%></td>
<td nowrap="nowrap"><div ><%=xlsrs1.Fields("�༶����").Value%></div></td>
<td nowrap="nowrap"><%=xlsrs1.Fields("�δ�").Value%></td>
<td nowrap="nowrap"><%=xlsrs1.Fields("�γ�����").Value%></td>
<td nowrap="nowrap"><%=xlsrs1.Fields("�Ͽ�ʱ��").Value%></td>
<td nowrap="nowrap"><%=xlsrs1.Fields("�¿�ʱ��").Value%></td>
<td nowrap="nowrap"><%=xlsrs1.Fields("��ʦ����").Value%></td>
<td nowrap="nowrap"><%=xlsrs1.Fields("��ѧ��").Value%></td>
<td nowrap="nowrap"><%=xlsrs1.Fields("����").Value%></td>
<td nowrap="nowrap"><%=xlsrs1.Fields("ĸ�׵绰").Value%></td>
<td nowrap="nowrap"><%=xlsrs1.Fields("���׵绰").Value%></td>
</tr>
<%
xlsrs1.MoveNext
Wend
%>
</table>
<%
xlsrs1.Close()
xlsconn1.Close()
Set xlsrs1 = Nothing
Set xlsconn1 = Nothing
%>      
</body>
</html>