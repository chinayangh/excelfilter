<%
	dim sql,filename,fs,myfile,x,conn,myconn1_Xsl,Path,rstData
	 
	Set conn = server.CreateObject("adodb.connection")
	'Set rstData =Server.CreateObject("Adodb.RecordSet")
	Set fs = server.CreateObject("scripting.filesystemobject")
	myconn1_Xsl="Provider=Microsoft.ACE.OLEDB.12.0;Data Source=c:\inetpub\wwwroot\excel\gz\file\book2.xls;Extended Properties=Excel 12.0"
	conn.open myconn1_Xsl


	'--�������������ɵ�EXCEL�ļ������µĴ��
	Path="file/"
	filename = Server.MapPath(path&"020data.xls")
	'--���ԭ����EXCEL�ļ����ڵĻ�ɾ����
	if fs.FileExists(filename) then
	fs.DeleteFile(filename)
	end if
	'--����EXCEL�ļ�
	set myfile = fs.CreateTextFile(filename,true)
	 
	sql = "select '����' as ��У,ѧԱuid,ѧԱ����,�༶����,�δ�,�γ�����,�Ͽ�ʱ��,�¿�ʱ��,��ʦ����,��ѧ��,����,ĸ�׵绰,���׵绰 from [Sheet0$]   where �༶���� not like '���ڰ�%' and �༶���� not like '���%' and �༶���� not like '������%' and �༶���� not like '�����%' and �༶���� not like '%����%' and �༶���� not like '%˫ʦ%' and �༶���� not like '%Сѧ���꼶����%' and �༶���� not like '%��ȤС��%' and �Ͽ�ʱ�� like '%10:30%' or �Ͽ�ʱ�� like '%08:00%' or �Ͽ�ʱ�� like '%08:30%' or �Ͽ�ʱ�� like '%12:00%' or �Ͽ�ʱ�� like '%13:20%' or �Ͽ�ʱ�� like '%14:50%' or �Ͽ�ʱ�� like '%15:50%' or �Ͽ�ʱ�� like '%18:00%' or �Ͽ�ʱ�� like '%18:30%' group by ѧԱuid,ѧԱ����,�༶����,�δ�,�γ�����,�Ͽ�ʱ��,�¿�ʱ��,��ʦ����,��ѧ��,����,ĸ�׵绰,���׵绰  order by ��ʦ���� "
	
	set rstData=conn.execute(sql)
	'rstData.open sql,conn,1,1
	
	if not rstData.EOF and not rstData.BOF then
	 
	dim strLine
	strLine=""
	For each x in rstData.fields
	strLine = strLine & x.name & chr(9)
	'response.write strLine
	Next
	 
	'--�����������д��EXCEL
	myfile.writeline strLine
	 
	Do while Not rstData.EOF
	strLine=""
	 
	for each x in rstData.Fields
	strLine = strLine & x.value & chr(9)
	next
	myfile.writeline strLine
	 
	rstData.MoveNext
	loop
	 
	end if
	'Response.Write("����EXCEL�ļ��ɹ������<a href=./file/020data.xls rel='external nofollow' target=_blank>����")
	Response.Write("ɸѡ�ɹ�����ʼת����ʽ���ذ�&nbsp&nbsp<a href=javascript:history.back(-1)>������һҳ</a>")
	rstData.Close
	set rstData = nothing
	Conn.Close
	Set Conn = nothing
	%>