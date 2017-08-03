<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Conn.asp"-->
<!--#include file="inc/md5.asp"-->
<!--#include file="inc/sha1.asp"-->
<!--#include file="inc/function.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>zmcms_管理页面</title>
</head>

<body>

<%
dim LoginName,LoginPassword,userName,Password,AdminPurview,Working,glyName,rs,sql
LoginName=trim(request.form("userName"))
LoginPassword=Md5(sha1(request.form("Password")&"zmcms"))
set rs = server.createobject("adodb.recordset")
sql="select * from admin where userName='"&LoginName&"'"
rs.open sql,conn,1,3

if rs.eof then
   response.write "管理员名称不正确，请重新输入"
   response.end
else
   userName=rs("userName")
   Password=rs("Password")
   AdminPurview=rs("AdminPurview")
   Working=rs("Working")
   glyName=rs("glyName")
end if

if LoginPassword<>Password then
   response.write "管理员密码不正确，请重新输入。"
   response.end
end if 


if TestCaptcha("ASPCAPTCHA", Request.Form("captchacode")) then
		else
	response.write "验证码输入不正确。"
   response.end
		end if

if working=0 then
   response.write "不能登录，此管理员帐号已被锁定。"
   response.end
end if 
 
if LoginName=userName and LoginPassword=Password then
   rs("LastLoginTime")=now()
   rs("LastLoginIP")=getip()
   rs.update
   rs.close
   set rs=nothing 
   session("userName")=userName
   session("glyName")=glyName
   session("password")=password
   session("AdminPurview")=AdminPurview
   session("LoginSystem")="Succeed"
   session.timeout=30
   '==================================
   dim LoginIP,LoginTime,LoginSoft
   LoginIP=getip()
   LoginSoft=Request.ServerVariables("Http_USER_AGENT")
   LoginTime=now()
   '====================================
   set rs = server.createobject("adodb.recordset")
   sql="select * from AdminLog"
   rs.open sql,conn,1,3
   rs.addnew
   rs("username")=username
   rs("glyname")=glyname
   rs("LoginIP")=LoginIP
   rs("LoginSoft")=LoginSoft
   rs("LoginTime")=LoginTime
   rs("Logcontent")="登录成功"
   rs.update
   rs.close
   set rs=nothing 
 
   '========================================
   response.Write "<img src=""images/loading.gif"" width=128 height=15 >"
%>

<script language="javascript">setTimeout("location.href='index.asp';",2000);</script> 
<%
end if
%>
</body>
</html>

