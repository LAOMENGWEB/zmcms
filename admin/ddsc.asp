<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="Inc/Function.asp"-->
<!--#include file="htmlconfig.asp"-->
<!--#include file="admin_html_function.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>zmcms_管理页面</title>
</head>

<body>
<%
If request.QueryString("ID") = "" Or Not IsNumeric(request.QueryString("ID")) Then
    response.Write "<script language=javascript>alert(""非法操作！"");window.history.back();</script>"
    response.End
End If

dim id
id=request.QueryString("id")
conn.Execute ( "delete from orderdd where orderddh ='"&id&"'" )
conn.Execute ( "delete from ordercp where orderddh ='"&id&"'" )
call addlog("删除订单")
response.Write "<script language=javascript>alert(""删除成功!"");window.location.href=""order.asp""</script>"

%>
</body>
</html>
