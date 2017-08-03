<!--#include file="../../inc/sub.asp"-->
<%
Response.Charset="utf-8"
if request.QueryString("Action")="Out" then
session.contents.remove "MemName"
session.contents.remove "GroupID"
session.contents.remove "GroupLevel"
session.contents.remove "MemLogin"
session.contents.remove "bangding"
WriteMsg 2,""&word(468)&"","member/memberuserlogin/?m="&request.QueryString("m")&"&Language="&Language&"","right"
ELSE
Public ErrMsg(5)
ErrMsg(0)=""&word(469)&""
ErrMsg(1)=""&word(470)&""
ErrMsg(2)=""&word(471)&""
ErrMsg(3)=""&word(472)&""
ErrMsg(4)=""&word(473)&""
LoginName=trim(request.form("LoginName"))
LoginPassword=Md5(request.form("LoginPassword"))
set rs = server.createobject("adodb.recordset")
sql="select * from Members where MemName='"&LoginName&"' and lang='"&language&"'"
rs.open sql,conn,1,3
if rs.eof then
WriteMsg 1,ErrMsg(0),"","wrong"
response.end
else
MemName=rs("MemName")
Password=rs("Password")
GroupID=rs("GroupID")
GroupName=rs("GroupName")
Working=rs("Working")
sh=rs("sh")
end if
if LoginPassword<>Password then
WriteMsg 1,ErrMsg(1),"","wrong"
response.end
end if 
if TestCaptcha("ASPCAPTCHA", Request.Form("captchacode")) then
session("checkcode")=""
else
WriteMsg 1,ErrMsg(4),"","wrong"
response.end
end if
if Working=0 then
WriteMsg 1,ErrMsg(2),"","wrong"
response.end
end if 
if hysh=1 then
if sh=0 then
WriteMsg 1,ErrMsg(3),"","wrong"
response.end
end if
end if
if UCase(LoginName)=UCase(MemName) and LoginPassword=Password then
rs("LastLoginTime")=now()
rs("LastLoginIP")=getip()
rs("LoginTimes")=rs("LoginTimes")+1
rs.update
rs.close
set rs=nothing
session("MemName")=MemName
session("GroupID")=GroupID
session("password")=password
session("working")=working
session("sh")=sh
   '===========
set rs = server.createobject("adodb.recordset")
sql="select * from hyGroup where GroupID='"&GroupID&"' and lang='"&language&"'"
rs.open sql,conn,1,1
session("GroupLevel")=rs("GroupLevel")
rs.close
set rs=nothing
  '===========
session("MemLogin")="Succeed"
session.timeout=60
WriteMsg 2,""&word(474)&"","member/membercenter/?m="&request.QueryString("m")&"&Language="&Language&"","right"
response.end
end if
end if
%>