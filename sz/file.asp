<%OPTION EXPLICIT%>
<%Server.ScriptTimeOut=5000%>
<!--#include FILE="upload_5xsoft.inc"-->
<%
'If IsEmpty(Session("username")) Then
'	response.write "<script>window.location.href='index.html'</script>"
'end if
dim upload,file,formName,formPath,fs,fs1
set upload=new upload_5xsoft ''�����ϴ�����
Set fs=Server.CreateObject("Scripting.FileSystemObject")
'Set fs1=Server.CreateObject("Scripting.FileSystemObject")

formPath="file/"'·��

for each formName in upload.objFile ''�г������ϴ��˵��ļ�
	set file=upload.file(formName)  ''����һ���ļ�����
	if file.FileSize>0 then         ''��� FileSize > 0 ˵�����ļ�����
		file.SaveAs Server.mappath(formPath&file.FileName)   ''�����ļ�
		session("fname")=file.FileName
		response.write session("fname")
		'fs.CopyFile "c:\inetpub\wwwroot\kdinfo.mdb","c:\inetpub\wwwroot\kdinfo.mdb."&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now),True
		fs.CopyFile "c:\inetpub\wwwroot\excel\sz\file\"&file.FileName,"c:\inetpub\wwwroot\excel\sz\file\book2.xls",True    '��fs��CopyFile���������ļ�

	end if
	set file=nothing
next

set upload=nothing  ''ɾ���˶���
%>
<script>window.parent.Finish("�ϴ��ļ��ɹ���");</script>