
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=title%>-<%=tw(70)%>-<%=tw(96)%></title>
<meta name="keywords" content="<%=gjc%>" />
<meta name="description" content="<%=ms%>" />
<link href="<%=template%>css/style.css" rel="stylesheet" type="text/css">
<!--#include file="../include/js.asp"-->
</head>
<body >

  
		
 			

<!--#include file="../top.asp"-->
<div class="wrap">
<div class="main">
 <!--左侧-->
<div class="fyLeft">
<dl class="l_pro">

<dt>-<%=tw(71)%></dt>

<dd><a href="<%=sitepath%>member/memberinfo/?m=<%=request.QueryString("m")%>&language=<%=language%>"><%=tw(72)%></a></dd>
<dd><a href="<%=sitepath%>member/memberguest/?m=<%=request.QueryString("m")%>&language=<%=language%>"><%=tw(73)%></a></dd>
<dd><a href="<%=sitepath%>member/memberorder/?m=<%=request.QueryString("m")%>&language=<%=language%>"><%=tw(74)%></a></dd>
<dd><a href="<%=sitepath%>member/memberjob/?m=<%=request.QueryString("m")%>&language=<%=language%>"><%=tw(75)%></a></dd>
<dd><a href="<%=sitepath%>inc/MemberLogin/?Action=Out&m=<%=request.QueryString("m")%>&language=<%=language%>"><%=tw(76)%></a></dd>





</dl>
<dl class="l_content">
<dt><%=tw(4)%>：<%=lxdz%></dt>
</dl>
<dl class="l_content">
<dt><%=tw(5)%>：<%=sj%></dt>
</dl>
<dl class="l_content">
<dt><%=tw(6)%>：<%=cz%></dt>
</dl>
<dl class="l_content">
<dt><%=tw(7)%>：<%=lxr%></dt>
</dl>

<div class="contact"> 
<a></a><a></a><a></a></div>
</div>

<!--右侧-->
<div class="fyRight">
<div class="title"> <span class="fl"><%=tw(70)%></span>
<div class="fr"><%=tw(10)%>：<%=tw(70)%>-<%=tw(96)%></div>
</div>


<div class="mainRightMain">
<!--会员中心 找回密码开始-->
<%
dim rs,sql,NewPassword
if request.QueryString("Step")="" then
%>
<form action="<%=sitepath%>member/MemberGetPass/?Step=EnterNameMail&m=<%=request.QueryString("m")%>&Language=<%=Language%>" method="post" class="checkform">
<table  class="membergetpass">
<tr>
<td class="zkd"><%=word(172)%></td>
<td class="ckd"><input name="MemName" type="text"  id="MemName"  class="inputxt" ajaxurl="<%=sitepath%>ajax/getpass.asp?language=<%=language%>"  datatype="*6-16" nullmsg="<%=word(173)%>" errormsg="<%=word(174)%>"/><span class="fh">*</span></td>
<td class="ykd"><div class="Validform_checktip"><%=word(175)%></div> </td>
</tr>
<tr>
<td class="zkd" > <%=word(176)%> </td>
<td class="ckd"><input name="Email" type="text"  id="Email"  class="inputxt" ajaxurl="<%=sitepath%>ajax/getpass.asp?language=<%=language%>" datatype="e" errormsg="<%=word(177)%>" nullmsg="<%=word(178)%>"/><span class="fh">*</span>
</td>
<td class="ykd"><div class="Validform_checktip"><%=word(179)%></div> </td>
 </tr>
<tr>
<td colspan="3" class="last">
<input name="Next" type="submit" value=" <%=word(180)%> " class="tjan"/>
</td>
</tr>
</table>
</form>
<%
elseif request.QueryString("Step")="EnterNameMail" then
set rs = server.createobject("adodb.recordset")
sql="select * from Members where MemName='"&request.form("MemName")&"' and Email='"&request.form("Email")&"' and Lang='"&Language&"'"
rs.open sql,conn,1,1
if rs.eof then
writemsg1 1,""&word(181)&"","","wrong"
end if
%>
<form action="<%=sitepath%>member/MemberGetPass/?Step=EnterAnswer&ID=<%=rs("ID")%>&m=<%=request.QueryString("m")%>&Language=<%=Language%>" method="post" class="checkform">
<table  class="membergetpass">
<tr>
<td class="zkd"><%=word(182)%></td>
<td class="ckd"><%=rs("MemName")%></td>
<td class="ykd"></td>
</tr>
<tr>
<td class="zkd"><%=word(183)%></td>
<td class="ckd"> <%=rs("Email")%></td>
<td class="ykd"></td>
</tr>
<tr>
<td class="zkd"><%=word(184)%></td>
<td class="ckd">
<%
if rs("Question")=1 then
response.write(""&word(138)&"")
end if
if rs("Question")=2 then
response.write(""&word(139)&"")
end if
if rs("Question")=3 then
response.write(""&word(140)&"")
end if
if rs("Question")=4 then
response.write(""&word(141)&"")
end if
if rs("Question")=5 then
response.write(""&word(142)&"")
end if
if rs("Question")=6 then
response.write(""&word(143)&"")
end if
if rs("Question")=7 then
response.write(""&word(144)&"")
end if
%>
</td>
<td class="ykd"></td>
</tr>
<tr>
<td class="zkd"><%=word(185)%></td>
<td class="ckd"><input name="Answer" type="text"   class="inputxt" ajaxurl="<%=sitepath%>ajax/getpass.asp?language=<%=language%>" datatype="*" nullmsg="<%=word(186)%>"/><span class="fh">*</span></td>
<td class="ykd"><div class="Validform_checktip"><%=word(187)%></div> </td>
</tr>
<tr>
<td colspan="3" style="border-right:none;border-bottom:none">
<input name="Next" type="submit" value=" <%=word(180)%> " class="tjan"/>
</td>
</tr>
</table>
</form>
<%
elseif request.QueryString("Step")="EnterAnswer" then
set rs = server.createobject("adodb.recordset")
sql="select * from Members where ID="&request.QueryString("ID")
rs.open sql,conn,1,3
randomize timer
NewPassword=Int(899999*Rnd() +100000)
if maillock=1 and getpass1=1 then
SendStat   =   Jmail(""&rs("Email")&"",""&word(188)&""&mc&""&word(189)&"",""&word(190)&"<br> "&word(191)&""&rs("MemName")&"<br>"&word(192)&""&NewPassword&"<br /><br />"&word(193)&"<br /><br />==============================<br>"&mc&"<br>"&word(194)&""&ym&"<br>"&word(195)&""&yx&"<br>"&word(196)&""&sj&"<br>"&word(197)&""&cz&"<br>"&word(198)&""&lxdz&"","GB2312","text/html")
end if
rs("Password")=MD5(NewPassword)
rs.update
rs.close
set rs=nothing 
if maillock=1 and getpass1=1 then
writemsg1 2,""&word(199)&"","","right"
else
response.write(""&word(515)&"【"&NewPassword&"】 "&word(516)&"")
end if

end if
%>


<!--会员中心 找回密码结束-->
</div>
</div>

</div>
</div>




<!--#include file="../bottom.asp"-->
</body>
</html>