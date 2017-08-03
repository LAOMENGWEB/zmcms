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
set rs=server.CreateObject("adodb.recordset")
sql = "select * from freeform where id = "&id&""
rs.open sql,conn,1,3
FormTableName=rs("FormTableName")
rs.Delete
strsql = "drop table "&FormTableName&""
conn.execute(strsql)
rs.close
set rs=nothing
conn.Execute ( "delete from FreeField where ChannelId="&id&"" )
call addlog("删除自定义表单")
response.Write "<script language=javascript>alert(""删除成功！"");window.location.href=""bd.asp""</script>"

%>
</body>
</html>
