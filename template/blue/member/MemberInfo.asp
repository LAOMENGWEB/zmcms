
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=mc%> >> -<%=tw(70)%>--<%=tw(97)%></title>
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

<dd><a href="<%=sitepath%>member/memberinfo/?m=<%=request.QueryString("m")%>"  <% if request.QueryString("m")=8 then response.write("class=""select""") end if%>><%=tw(72)%></a></dd>
<dd><a href="<%=sitepath%>member/memberguest/?m=<%=request.QueryString("m")%>&language=<%=language%>"><%=tw(73)%></a></dd>
<dd><a href="<%=sitepath%>member/memberorder/?m=<%=request.QueryString("m")%>&language=<%=language%>"><%=tw(74)%></a></dd>
<dd><a href="<%=sitepath%>member/memberjob/?m=<%=request.QueryString("m")%>&language=<%=language%>"><%=tw(75)%></a></dd>
<dd><a href="<%=sitepath%>inc/MemberLogin/?Action=Out&m=<%=request.QueryString("m")%>&language=<%=language%>"><%=tw(76)%></a></dd>





</dl>
<dl class="l_content">
<dt>-<%=tw(4)%>：<%=lxdz%></dt>
</dl>
<dl class="l_content">
<dt>-<%=tw(5)%>：<%=sj%></dt>
</dl>
<dl class="l_content">
<dt>-<%=tw(6)%>：<%=cz%></dt>
</dl>
<dl class="l_content">
<dt>-<%=tw(7)%>：<%=lxr%></dt>
</dl>

</div>

<!--右侧-->
<div class="fyRight">
<div class="title"> <span class="fl"><%=tw(70)%></span>
<div class="fr"><%=tw(10)%>：<%=tw(70)%> >> <%=tw(97)%></div>
</div>


<div class="mainRightMain">
<!--会员中心-会员资料代码开始-->
<%
if not (session("MemName")<>"" and session("MemLogin")="Succeed") then 
writemsg1 2,""&word(95)&"","member/memberuserlogin/?m="&request.QueryString("m")&"&language="&language&"","wrong"
response.end
end if
%>
<form action="<%=sitepath%>inc/MemberInfo/?ID=<%=MemIDmemberinfo%>&m=<%=request.QueryString("m")%>&language=<%=language%>" method="post" class="checkform">
<table  cellpadding="0" cellspacing="0" class="memberinfo">
<tr>
<td colspan="3" class="text"><%=word(96)%></td>
</tr>
<tr>
<td class="zkd"><%=word(97)%></td>
<td colspan="2" style="border-right:none"><div class="jlz"><% =mMemNamememberinfo %></div></td>
</tr>
<tr>
<td class="zkd"><%=word(98)%></td>
<td colspan="2" style="border-right:none"><div class="jlz"><font color="#ff6600"><% =mGroupIdNamememberinfo %></font></div></td>
</tr>
<tr>
<td class="zkd"><%=word(99)%></td>
<td colspan="2" style="border-right:none"><div class="jlz"><input name="RealName" type="text" id="RealName" value="<%=mRealNamememberinfo%>" size="20" maxlength="50" class="wbk"/></div></td>
</tr>
<tr>
<td class="zkd"><%=word(100)%></td>
<td colspan="2" style="border-right:none"><div class="jlz">
<input name="Sex" type="radio"  value="<%=word(512)%>" <%if mSexmemberinfo=""&word(512)&"" then response.write ("checked")%>  style="border:none"/>&nbsp;<%=word(512)%>
&nbsp;&nbsp;<input type="radio" name="Sex" value="<%=word(513)%>" <%if mSexmemberinfo=""&word(513)&"" then response.write ("checked")%> style="border:none"/>&nbsp;<%=word(513)%>
</div>						
</td>
</tr>
<tr>
<td class="zkd"><%=word(101)%></td>
<td colspan="2" style="border-right:none"><div class="jlz"><input name="Password" type="password" id="Password" size="20" maxlength="16" class="wbk"/> <%=word(514)%></div></td>
</tr>
<tr>
<td class="zkd"><%=word(102)%></td>
<td colspan="2" style="border-right:none"><div class="jlz"><input name="vPassword" type="password" id="vPassword" size="20" maxlength="16" class="wbk"/></div></td>
</tr>
<tr>
<td class="zkd"><%=word(103)%></td>
<td colspan="2" style="border-right:none">
<div class="jlz">


<%
if mQusetionmemberinfo=1 then
response.write(""&word(138)&"")
end if
if mQusetionmemberinfo=2 then
response.write(""&word(139)&"")
end if
if mQusetionmemberinfo=3 then
response.write(""&word(140)&"")
end if
if mQusetionmemberinfo=4 then
response.write(""&word(141)&"")
end if
if mQusetionmemberinfo=5 then
response.write(""&word(142)&"")
end if
if mQusetionmemberinfo=6 then
response.write(""&word(143)&"")
end if
if mQusetionmemberinfo=7 then
response.write(""&word(144)&"")
end if
%>
</div></td>
</tr>
<tr>
<td class="zkd"><%=word(104)%></td>

<td colspan="2" style="border-right:none"><div class="jlz">********</div></td>
</tr>
<tr>
<td class="zkd"><%=word(105)%></td>
<td colspan="2" style="border-right:none"><div class="jlz"><input name="Company" type="text" id="Company" value="<% =mCompanymemberinfo %>" size="40" maxlength="100" class="wbk"/></div></td>
</tr>
<tr>
<td class="zkd"><%=word(106)%></td>
<td colspan="2" style="border-right:none"><div class="jlz"><input name="Address" type="text" id="Address" value="<% =mAddressmemberinfo %>" size="40" maxlength="100" class="wbk"/></div></td>
</tr>
<tr>
<td class="zkd"><%=word(107)%></td>
<td colspan="2" style="border-right:none"><div class="jlz"><input name="ZipCode" type="text" id="ZipCode" value="<% =mZipCodememberinfo %>" size="20" maxlength="20" class="wbk"/></div></td>
</tr>
<tr>
<td class="zkd"><%=word(108)%></td>
<td colspan="2" style="border-right:none"><div class="jlz"><input name="Fax" type="text" id="Fax" value="<% =mFaxmemberinfo %>" size="20" maxlength="50" class="wbk"/></div></td>
</tr>
<tr>
<td class="zkd"><%=word(109)%></td>
<td colspan="2" style="border-right:none"><div class="jlz"><input name="Mobile" type="text" id="Mobile" value="<% =mMobilememberinfo %>" size="20" maxlength="50" class="wbk"/></div></td>
</tr>
<tr>
<td class="zkd"><%=word(110)%></td>
<td colspan="2" style="border-right:none"><div class="jlz"><input name="Email" type="text" id="Email" value="<% =mEmailmemberinfo %>" size="30" maxlength="50" class="wbk"/></div></td>
</tr>
<tr>
<td class="zkd"><%=word(111)%></td>
<td width="101" style="border-right:none"><div class="jlz"><input name="captchacode" id=name="captchacode" type="text" size="6" maxlength="6" style="width:98px;" class="wbk"/></div></td>
<td width="753" style="border-right:none;padding-left:10px"><img id="imgCaptcha" src="<%=sitepath%>inc/captcha.asp" />&nbsp;<a href="###" onClick="RefreshImage('imgCaptcha')"><img src="<%=sitepath%>images/reload.png" width="20" height="20" /></a></td>
</tr>
<tr>
<td class="zkd"><%=word(112)%></td>
<td colspan="2" style="border-right:none"><div class="jlz"><% =mAddTimememberinfo %></div></td>
</tr>
<tr>
<td class="zkd"><%=word(113)%></td>
<td colspan="2" style="border-right:none"><div class="jlz"><% =mLoginTimesmemberinfo %></div></td>
</tr>
<tr>
<td class="zkd"><%=word(114)%></td>
<td colspan="2" style="border-right:none"><div class="jlz"><% =mLastLoginTimememberinfo %></div></td>
</tr>
<tr>
<td class="zkd"><%=word(115)%></td>
<td colspan="2" style="border-right:none"><div class="jlz"><% =mLastLoginIPmemberinfo %> (<%=gethttp("http://whois.pconline.com.cn/ip.jsp?ip="&mLastLoginIPmemberinfo&"","")%>)</div></td>
</tr>
<tr>
<td colspan="3" style="border-right:none; border-bottom:none"><input name="Submit" type="submit" value=" <%=word(116)%> " class="tjan" />						</td>
</tr>

</table>
</form>



<!--会员中心-会员资料代码结束-->	
</div>
</div>

</div>
</div>




<!--#include file="../bottom.asp"-->




</body>
</html>