﻿<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="Inc/Function.asp"-->
<!--#include file="admin_html_function.asp"-->
<!--#include file="htmlconfig.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>zmcms_管理页面</title>
</head>

<body>
<%
dim id
id=request.QueryString("id")
set rs=server.CreateObject("adodb.recordset")
sql = "select * from dy where id = "&id&""
rs.open sql,conn,1,3
IcoImage=rs("IcoImage")
SmallImage=rs("SmallImage")
zdfy=rs("zdfy")
pone=rs("zmone")
if rs.bof and rs.eof then
response.Write "<script language=javascript>alert(""不存在此数据！"");window.history.back();</script>"
response.End()
end if
dim picUrl
dim picUrlArray
dim x,y
picUrl = getPicUrl(rs("content"))
if picUrl <> "" then
picUrl = left(picUrl,len(picUrl)-1)
picUrlArray = split(picUrl,",")
for x = 0 to ubound(picUrlArray)
if instr(picUrlArray(x),"") > 0 then
call DeleteFile(picUrlArray(x))
end if
next
end if

dim picUrl1
dim picUrlArray1
dim x1,y1
picUrl1 = rs("ImagePath")
if picUrl1 <> "" then
picUrlArray1 = split(picUrl1,"|")
for x1 = 0 to ubound(picUrlArray1)
DeleteFile(picUrlArray1(x1))
next
end if


if Instr(rs("content"),"[zmcmsfy]")=0 then
call DeleteFile("../html/"&aboutname&""&Separated&""&pone&""&Separated&""&ID&""&Separated&"1.html")
else
arrContent=split(rs("content"),"[zmcmsfy]")
pages=Ubound(arrContent)+1
for i=1 to pages
call DeleteFile("../html/"&aboutname&""&Separated&""&pone&""&Separated&""&ID&""&Separated&""&i&".html")
next
end if


call deletefile1(SmallImage)
call deletefile(IcoImage)


rs.delete
rs.close
set rs=nothing
if yy=2 then
call htmlabout
call htmll("","","default"&lang&".html","default.asp","m=",mde,"language=",lang,"","","","","","","","")
end if
call addlog("删除"&danye&"信息")
response.Write "<script language=javascript>alert(""删除成功！"");window.location.href=""dy.asp""</script>"

%>
</body>
</html>
