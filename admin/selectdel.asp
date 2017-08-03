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
if Instr(session("AdminPurview"),"|36,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</script>"
  response.end
end if
'========判断是否具有管理权限
%>
<%
dim id
id=request.QueryString("id")
mode=request.QueryString("mode")
conn.Execute ( "delete from SelectName where s_id="&id&"" )
conn.Execute ( "delete from SelectValue where s_id="&id&"" )
if mode=0 then
call addlog("删除筛选属性")
end if
if mode=1 then
call addlog("删除规格属性")
end if
if mode=2 then
call addlog("删除商品属性")
end if
response.Write "<script language=javascript>alert(""删除成功！"");window.location.href=""select.asp?mode="&mode&"""</script>"


%>
</body>
</html>
