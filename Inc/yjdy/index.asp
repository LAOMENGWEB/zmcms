<!--#include file="../sub.asp"-->
<%
Response.Charset="utf-8"
mail=request.querystring("mail")
type1=request.querystring("do")
U=request.querystring("U")
if DateDiff("s",session("time"),now())<5 then  
WriteMsg 2,""&word(370)&"",""&u&"","wrong"
response.end
else  
session("time")=now()
end if
if chkpost=0 then
WriteMsg 1,""&word(491)&"","","wrong"
response.end
end if
if type1=1 then
exec="select * from demaill where emaill='"&mail&"'"
set rs=server.createobject("adodb.recordset")
rs.open exec,conn,1,3
if not  rs.eof then
writemsg 2,""&word(461)&"",""&u&"","wrong"
else
rs.addnew
rs("emaill")=mail
rs("adddate")=now()
rs("lang")=language
rs.update
memo=""&word(462)&""+mail+""&word(463)&""
writemsg 2,""&memo&"",""&u&"","right"
rs.close
set rs=nothing
end if 
else
set rs=server.createobject("adodb.recordset")
exec="select * from demaill where emaill='"&mail&"'"
rs.open exec,conn,1,3
if not rs.eof then
exe="delete * from demaill where emaill='"&mail&"'"
conn.execute exe
writemsg 2,""&word(464)&"",""&u&"","right"
else
writemsg 2,""&word(465)&"",""&u&"","right"
end if 
end if 
%>
