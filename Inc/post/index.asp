<!--#include file="../sub.asp"-->
<%
Response.Charset="utf-8"
if DateDiff("s",session("time"),now())<5 then  
WriteMsg 1,""&word(370)&"","","wrong"
response.end
else  
session("time")=now()
end if
if chkpost=0 then
WriteMsg 1,""&word(491)&"","","wrong"
response.end
end if
mMemID=request.QueryString("MemberID")
set rs=server.createobject("adodb.recordset")
sql="select * from ly "
if trim(request.form("qqh"))="1" then
qqh=1
else
qqh=0
end if
rs.open sql,conn,1,3
title=Trim(request.form("title")) 
phone=Trim(request.form("phone"))
emaill=Trim(request.form("emaill"))
qq=Trim(request.form("qq"))
content=Trim(request.form("content"))



if maillock=1 and newguest=1 then 
SendStat   =   Jmail(""&PostMail&"",""&word(406)&""&title&""&word(407)&"",""&word(408)&""&title&""&word(409)&"<br>"&word(410)&""&phone&"<br>"&word(411)&""&emaill&"<br>"&word(412)&""&qq&"<br>"&word(413)&""&content&"","GB2312","text/html")   
end if
rs.addnew
rs("title")=title
rs("emaill")=emaill
rs("phone")=phone
rs("sj")=now()
rs("content")=content
rs("qq")=qq
rs("lyip")=getip()
rs("qqh")=qqh
rs("MemID")=mMemID
rs("lang")=language
if ly=1 then
rs("sh")=0
end if
if ly=2 then
rs("sh")=1
end if
rs.update
rs.Close
Set rs = Nothing
if ly=1 then	
writemsg 2,""&word(443)&"",""&showguest&"/?m="&request.QueryString("m")&"&Language="&Language&"","right"
else
if maillock=1 and newguest=1 then
writemsg 2,""&word(444)&"",""&showguest&"/?m="&request.QueryString("m")&"&Language="&Language&"","right"
else
writemsg 2,""&word(445)&"",""&showguest&"/?m="&request.QueryString("m")&"&Language="&Language&"","right"
end if 
end if
%>