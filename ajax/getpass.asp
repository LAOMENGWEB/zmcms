
<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!-- #include file="../inc/sub.asp" --> 

<%
param=request.form("param")
name1=request.form("name")
if name1="MemName" then

Set rs = server.CreateObject("adodb.recordset")
rs.open "Select * From members Where MemName='" & param & "'",conn,1,3
if not (rs.bof and rs.EOF) then
response.write "{""info"":"""&word(362)&""",""status"":""y""}"
else
response.write "{""info"":"""&word(363)&""",""status"":""x""}"
end if
end if
if name1="Email" then
Set rs = server.CreateObject("adodb.recordset")
rs.open "Select * From members Where Email='" & param & "'",conn,1,3
if not (rs.bof and rs.EOF) then
response.write "{""info"":"""&word(362)&""",""status"":""y""}"
else
response.write "{""info"":"""&word(364)&""",""status"":""x""}"
end if
end if

if name1="Answer" then
Set rs = server.CreateObject("adodb.recordset")
rs.open "Select * From members Where Answer='" & param & "'",conn,1,3
if not (rs.bof and rs.EOF) then
response.write "{""info"":"""&word(362)&""",""status"":""y""}"
else
response.write "{""info"":"""&word(365)&""",""status"":""x""}"
end if
end if

%>
