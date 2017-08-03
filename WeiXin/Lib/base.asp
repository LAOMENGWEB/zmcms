<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
	session.codepage=65001
	response.charset="utf-8"
	server.scripttimeout=999999
	dim startime:startime=timer()
%>
<!--#include file="../Config.asp"-->
<!--#include file="db.asp"-->
<!--#include file="function.asp"-->
<!--#include file="weixin_class.asp"-->
<!--#include file="md5.asp"-->
<%
	dim sqltime:sqltime="now()"
	if not(datatype) then sqltime="GetDate()"
	Cls.db.dbopen()
%>