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
call DeleteFile("../default"&lang&".html")
call DeleteFile("../"&iinfo&"/"&Newclass&".html")
call DeleteFile("../"&product&"/"&Productclass&".html")
call DeleteFile("../"&photo&"/"&alclass&".html")
call DeleteFile("../"&video&"/"&tvclass&".html")
call DeleteFile("../"&job&"/"&zpclass&".html")
call DeleteFile("../"&down&"/"&downclass&".html")
response.Write "<script language=javascript>alert(""恭喜您删除成功！"");window.history.back(-2);</script>"

%>
</body>
</html>
