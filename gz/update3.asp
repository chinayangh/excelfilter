<%
	dim sql,filename,fs,myfile,x,conn,myconn1_Xsl,Path,rstData
	 
	Set conn = server.CreateObject("adodb.connection")
	Set rstData =Server.CreateObject("Adodb.RecordSet")
	Set fs = server.CreateObject("scripting.filesystemobject")
	myconn1_Xsl="Provider=Microsoft.ACE.OLEDB.12.0;Data Source=c:\inetpub\wwwroot\excel\gz\file\book2.xls;Extended Properties=Excel 12.0"
	conn.open myconn1_Xsl


	'--假设你想让生成的EXCEL文件做如下的存放
	Path="file/"
	filename = Server.MapPath(path&"020.xls")
	'--如果原来的EXCEL文件存在的话删除它
	if fs.FileExists(filename) then
	fs.DeleteFile(filename)
	end if
	'--创建EXCEL文件
	set myfile = fs.CreateTextFile(filename,true)

	 
	sql = "select '广州' as 分校,学员uid,学员名称,班级名称,课次,课程日期,上课时间,下课时间,教师名称,count(教师名称) as num,教学点,教室,电话1,电话2 from [Sheet0$]   where 班级名称 not like '短期班%' and 班级名称 not like '活动类%' and 班级名称 not like 'X试类%' and 班级名称 not like '诊X类%' and 班级名称 not like '%高中%' and 班级名称 not like '%双X%' and 班级名称 not like '%小学五年级X%' and 班级名称 not like '%兴趣小组%' and (上课时间 like '%10:30%' or 上课时间 like '%08:00%' or 上课时间 like '%08:30%' or 上课时间 like '%12:00%' or 上课时间 like '%13:20%' or 上课时间 like '%14:50%' or 上课时间 like '%15:50%' or 上课时间 like '%18:00%' or 上课时间 like '%18:30%') group by 学员uid,学员名称,班级名称,课次,课程日期,上课时间,下课时间,教师名称,教学点,教室,电话1,电话2 order by 教师名称 "
	
	
	rstData.open sql,conn,1,1
	
	if not rstData.EOF  then
	 
	dim strLine
	strLine=""
	For each x in rstData.fields
	strLine = strLine & x.name & chr(9)
	'response.write strLine
	Next
	 
	'--将表的列名先写入EXCEL
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
	Response.Write("生成EXCEL文件成功，点击<a href=./file/020.xls rel='external nofollow' target=_blank>下载")
	rstData.Close
	set rstData = nothing
	Conn.Close
	Set Conn = nothing
	%>
