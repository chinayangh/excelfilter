<%OPTION EXPLICIT%>
<%Server.ScriptTimeOut=5000%>
<!--#include FILE="upload_5xsoft.inc"-->
<%
'If IsEmpty(Session("username")) Then
'	response.write "<script>window.location.href='index.html'</script>"
'end if
dim upload,file,formName,formPath,fs,fs1
set upload=new upload_5xsoft ''建立上传对象
Set fs=Server.CreateObject("Scripting.FileSystemObject")
'Set fs1=Server.CreateObject("Scripting.FileSystemObject")

formPath="file/"'路径

for each formName in upload.objFile ''列出所有上传了的文件
	set file=upload.file(formName)  ''生成一个文件对象
	if file.FileSize>0 then         ''如果 FileSize > 0 说明有文件数据
		file.SaveAs Server.mappath(formPath&file.FileName)   ''保存文件
		session("fname")=file.FileName
		response.write session("fname")
		'fs.CopyFile "c:\inetpub\wwwroot\kdinfo.mdb","c:\inetpub\wwwroot\kdinfo.mdb."&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now),True
		fs.CopyFile "c:\inetpub\wwwroot\excel\sz\file\"&file.FileName,"c:\inetpub\wwwroot\excel\sz\file\book2.xls",True    '用fs的CopyFile方法复制文件

	end if
	set file=nothing
next

set upload=nothing  ''删除此对象
%>
<script>window.parent.Finish("上传文件成功！");</script>