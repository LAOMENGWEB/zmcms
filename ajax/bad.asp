
<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../inc/sub.asp"-->

<%
param=request.form("param")
name1=request.form("name")

if not ChkSB(param) then
response.write "{""info"":"""&word(361)&""",""status"":""x""}"
else
response.write "{""info"":"""&word(362)&""",""status"":""y""}"
response.End
End If



%>
