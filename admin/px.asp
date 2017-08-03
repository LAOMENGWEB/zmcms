<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>zmcms_管理页面</title>
</head>

<body>
<%
id=request("id")
for i=1 to request.form("id").count
if request.form("px")(i)<=0 then
px=0
else
px=request.form("px")(i)
end if
set rs_s=server.CreateObject("adodb.recordset")
conn.execute("update Definitons set px="&px&" where id="&request.form("id")(i))
next
response.Write "<script language=javascript>window.history.back(-1);</script>"
%>
</body>
</html>
