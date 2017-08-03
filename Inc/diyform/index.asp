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
id=Request.QueryString("id")
url=Request.QueryString("url")
set rsObj1=server.CreateObject("adodb.recordset")
Sqlp ="select * from FreeForm Where id="&id&""     
rsObj1.open sqlp,conn,1,1 


Set rs = server.CreateObject("adodb.recordset")
rs.open "Select * From "&rsObj1("FormTableName")&"" ,conn,1,3
rs.addnew

set rsObj=server.CreateObject("adodb.recordset")
Sqlp1 ="select * from FreeField where ChannelId="&id&""
rsObj.open sqlp1,conn,1,1

do while not rsObj.eof
zdyzd1=rsObj("FieldName")
zdyzd2=request.form(""&zdyzd1&"")
rs(""&zdyzd1&"")=zdyzd2	
rsObj.movenext
loop


rs("Ip")=getip()

if rsObj1("fsh")=1 then
rs("sh")=0
end if
if rsObj1("fsh")=2 then
rs("sh")=1
end if
rs("adddate")=now()
rs("lang")=language
rs.Update
writemsg 2,""&rsObj1("SuccessMsg")&"",""&url&"","right"
rsObj.close
set rsObj=nothing
rsObj1.close
set rsObj1=nothing


%>