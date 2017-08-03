
<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../inc/sub.asp"-->

<%
param=request.form("param")
name1=request.form("name")

if TestCaptcha("ASPCAPTCHA",param) then
response.write "{""info"":"""&word(362)&""",""status"":""y""}"
		else
response.write "{""info"":"""&word(368)&""",""status"":""x""}"
		end if



%>
