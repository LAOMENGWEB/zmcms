<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
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
diyform=request.QueryString("form")
id=request.QueryString("id")
id1=request.QueryString("id1")
set rs=server.CreateObject("adodb.recordset")
sql = "select * from FreeField where id = "&id&""
rs.open sql,conn,1,3
FieldName=rs("FieldName")
rs.Delete
DelColumn ""&diyform&"",""&FieldName&""
rs.close
set rs=nothing
call addlog("删除字段")
response.Write "<script language=javascript>alert(""删除成功！"");window.location.href=""zd.asp?form="&diyform&"&id="&id1&"""</script>"

%>
</body>
</html>
