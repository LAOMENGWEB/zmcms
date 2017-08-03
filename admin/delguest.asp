<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="inc/head.asp"-->
<!--#include file="Inc/Function.asp"-->
<%
if Instr(session("AdminPurview"),"|81,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</script>"
  response.end
end if
'========判断是否具有管理权限
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>zmcms_管理页面</title>


</head>

<body>

<%

totalrec=Conn.Execute("Select count(*) from ly")(0)
totalpage=int(totalrec/lyinfo)       
If (totalpage * lyinfo)<totalrec Then
totalpage=totalpage+1
End If
if totalpage<=1 then
call DeleteFile(""&showguest&"/"&guestname&""&Separated&"1.html")
else
for i=1 to totalpage
call DeleteFile(""&showguest&"/"&guestname&""&Separated&""&i&".html")
next
end if


call addlog("删除html")
response.Write "<script language=javascript>alert(""删除成功"");window.location.href=""right.asp""</script>"
response.End

%>
</body>
</html>
