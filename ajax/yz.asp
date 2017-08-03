
<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!-- #include file="../inc/sub.asp" --> 

<%
param=request.form("param")
name1=request.form("name")
if name1="MemName" then
if not pbhy(param) then
response.write "{""info"":"""&word(366)&""",""status"":""x""}"
else
Set rs = server.CreateObject("adodb.recordset")
rs.open "Select * From members Where MemName='" & param & "'",conn,1,3
if not (rs.bof and rs.EOF) then
response.write "{""info"":"""&word(366)&""",""status"":""x""}"
else
response.write "{""info"":"""&word(362)&""",""status"":""y""}"
end if
end if
end if
if name1="Email" then
Set rs = server.CreateObject("adodb.recordset")
rs.open "Select * From members Where Email='" & param & "'",conn,1,3
if not (rs.bof and rs.EOF) then
response.write "{""info"":"""&word(367)&""",""status"":""x""}"
else
response.write "{""info"":"""&word(362)&""",""status"":""y""}"
end if
end if
%>
