<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="qqconnect.asp"-->
<%
'解绑成功
if request.QueryString("act")="jb" then
Conn.execute("update members set qqopenapi='' where MemName='" & session("MemName") & "'")
WriteMsg 2,""&word(492)&"","member/membercenter/?m="&m6&"&Language="&Language&"","right"
response.end
session("bangding")=""
end if
'在登陆情况下绑定QQ
if Session("bangding")=1 then
Conn.execute("update members set qqopenapi='"&Session("Openid")&"' where MemName='" & session("MemName") & "'")
WriteMsg 2,""&word(487)&"","member/membercenter/?m="&m6&"&Language="&Language&"","right"
response.end
session("bangding")=""
end if
'第一次QQ登陆下绑定用户名
if request.QueryString("act")="bd" then
MemName=request.form("MemName")
qqapi=request.form("qqapi")
loginpass=md5(request.form("loginpass"))
Conn.execute("update members set qqopenapi='" & qqapi & "' where MemName='" & MemName & "' and Lang='"&Language&"'")
set rs2=server.createobject("adodb.recordset")
sql2="select MemName,GroupID,Password,Working,sh,LastLoginTime,LastLoginIP,LoginTimes from members where MemName='" & MemName & "' and qqopenapi='" & qqapi & "' and Lang='"&Language&"' "  
rs2.open sql2,conn,1,1
if not rs2.eof then
if loginpass<>rs2(2) then
WriteMsg 1,""&word(488)&"","","wrong"
else
if hysh=1 then
if rs2(4)=0 then
WriteMsg 2,""&word(489)&"","member/membercenter/?m="&m6&"&Language="&Language&"","wrong"
response.end
end if
end if


if rs2(3)=0 then
Conn.execute("update members set qqopenapi='' where MemName='" & MemName & "' and Lang='"&Language&"'")

WriteMsg 2,""&word(490)&"","member/membercenter/?m="&m6&"&Language="&Language&"","wrong"
response.end

else
session("MemName") = rs2(0)
session("GroupID") = rs2(1)
session("MemLogin")="Succeed"
if sqlno=1 then
Conn.execute("update members set LastLoginTime = now() where MemName='" & MemName & "' and Lang='"&Language&"'")
else
Conn.execute("update members set LastLoginTime = getdate() where MemName='" & MemName & "' and Lang='"&Language&"'")
end  if
Conn.execute("update members set LastLoginIP = '"&getip&"' where MemName='" & MemName & "' and Lang='"&Language&"'")
Conn.execute("update members set LoginTimes = LoginTimes + 1 where MemName='" & MemName & "' and Lang='"&Language&"'")
WriteMsg 2,""&word(487)&"","member/membercenter/?m="&m6&"&Language="&Language&"","right"
end if
end if
end if
end if
'第一次登陆QQ 认证
'判断登陆账号是否绑定！
SET qc = New QqConnet
CheckLogin=qc.CheckLogin()
If CheckLogin=False Then
else
'Response.Write(Session("State"))
Session("Access_Token")=qc.GetAccess_Token()
Session("Openid")=qc.Getopenid()
'Response.Write Session("Openid")
set rs=server.createobject("adodb.recordset")
sql="select MemName,GroupID,Working,sh,LastLoginTime,LastLoginIP,LoginTimes from members where qqopenapi='"&Session("Openid")&"' and Lang='"&Language&"'"  
rs.open sql,conn,1,1
if not rs.eof then
if hysh=1 then
if rs(3)=0 then
WriteMsg 2,""&word(472)&"","member/membercenter/?m="&m6&"&Language="&Language&"","wrong"
response.end
end if
end if
if rs(2)=0 then
WriteMsg 2,""&word(471)&"","member/membercenter/?m="&m6&"&Language="&Language&"","wrong"
response.end
else
session("MemName") = rs(0)
session("GroupID") = rs(1)
session("MemLogin")="Succeed"
if sqlno=1 then
Conn.execute("update members set LastLoginTime = now() where MemName='"&Session("MemName")&"' and Lang='"&Language&"'")
else
Conn.execute("update members set LastLoginTime = getdate() where MemName='"&Session("MemName")&"' and Lang='"&Language&"'")
end if
Conn.execute("update members set LastLoginIP = '"&getip&"' where MemName='"&Session("MemName")&"' and Lang='"&Language&"'")
Conn.execute("update members set LoginTimes = LoginTimes + 1 where MemName='"&Session("MemName")&"' and Lang='"&Language&"'")
WriteMsg 2,""&word(474)&"","member/membercenter/?m="&m6&"&Language="&Language&"","right"
response.end
end if
else
%>


<%
echo ob_get_contents(""&template&"api/user.asp")
%>

<%
end if
end if
%>