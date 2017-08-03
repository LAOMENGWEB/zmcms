<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="Inc/Function.asp"-->


<%

if Instr(session("AdminPurview"),"|614,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</script>"
  response.end
end if




%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>zmcms_管理页面</title>
</head>

<body>
<%
t=request.QueryString("t")
dim id
id=request.QueryString("id")
set rs=server.CreateObject("adodb.recordset")
sql = "select * from templateTags where id = "&id&""
rs.open sql,conn3,1,3
if rs.bof and rs.eof then
response.Write "<script language=javascript>alert(""不存在此数据！"");window.history.back();</script>"
response.End()
else
rs.delete
if t=1 then
response.Write "<script language=javascript>alert(""删除成功！"");window.location.href=""tagstyle.asp?t=1""</script>"
end if
if t=2 then
response.Write "<script language=javascript>alert(""删除成功！"");window.location.href=""tagstyle.asp?t=2""</script>"
end if

end if
rs.close

%>
</body>
</html>