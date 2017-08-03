<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="Inc/Function.asp"-->
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
l=request.QueryString("l")
If l="q" then
set rs=server.CreateObject("adodb.recordset")
sql = "select * from lang where qt = 1"
rs.open sql,conn,1,3
rs("qt")=0
rs.update
rs.close
set rs=nothing
set rs=server.CreateObject("adodb.recordset")
sql = "select * from lang where id = "&id&""
rs.open sql,conn,1,3
rs("qt")=1
rs.update
rs.close
set rs=Nothing
call addlog("设为前台默认语言")
End if
If l="h" then
set rs=server.CreateObject("adodb.recordset")
sql = "select * from lang where ht = 1"
rs.open sql,conn,1,3
rs("ht")=0
rs.update
rs.close
set rs=nothing
set rs=server.CreateObject("adodb.recordset")
sql = "select * from lang where id = "&id&""
rs.open sql,conn,1,3
rs("ht")=1
rs.update
rs.close
set rs=Nothing
call addlog("设为后台默认语言")
End if
response.Write "<script language=javascript>alert(""操作成功！"");window.location.href=""lang.asp""</script>"



%>
</body>
</html>
