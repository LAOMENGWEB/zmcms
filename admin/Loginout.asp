<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<%
call addlog("退出登录")
session("userName")=""
Response.Redirect "Login.asp"
%>
