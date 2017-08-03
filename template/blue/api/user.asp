

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=title%>-<%=tw(70)%></title>
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

<dt><%=tw(71)%></dt>

<dd><a href="<%=sitepath%>member/memberinfo/?m=<%=m2%>&language=<%=lanyybs%>"><%=tw(72)%></a></dd>
<dd><a href="<%=sitepath%>member/memberguest/?m=<%=m2%>&language=<%=lanyybs%>"><%=tw(73)%></a></dd>
<dd><a href="<%=sitepath%>member/memberorder/?m=<%=m2%>&language=<%=lanyybs%>"><%=tw(74)%></a></dd>
<dd><a href="<%=sitepath%>member/memberjob/?m=<%=m2%>&language=<%=lanyybs%>"><%=tw(75)%></a></dd>
<dd><a href="<%=sitepath%>inc/MemberLogin/?Action=Out&m=<%=m2%>&language=<%=lanyybs%>"><%=tw(76)%></a></dd>
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
</div>

<!--右侧-->
<div class="fyRight">
<div class="title"> <span class="fl"><%=tw(70)%></span>
<div class="fr"><%=tw(10)%>：<%=tw(70)%></div>
</div>

<div class="mainRightMain">

<%=tw(284)%>

<div class="fade">
<ul class="tabs">
<li><a href="#con1"><span><%=tw(285)%></span></a></li>

<li><a href='#con2'><span><%=tw(286)%></span></a></li>

</ul>
<div class="items">
<div id="con1"><form id="form1" name="form1" method="post" action="?act=bd&language=<%=lanyybs%>" class="checkform">

<table  class="hydl">
  <tr>
    <td class="zkd" align="right"><%=tw(287)%>：</td>
    <td style="border-right:none"><input name="MemName" type="text" class="wbk"  datatype="*" nullmsg="<%=word(120)%>" /><td class="ykd"><div class="Validform_checktip"></div>
	
	
	<input name="qqapi" type="hidden" value="<%=Session("Openid")%>" />
	</td>
  </tr>
   <tr>
    <td class="zkd" align="right"><%=tw(288)%>：</td>
    <td style="border-right:none"><input name="loginpass" type="password" class="wbk"  datatype="*" nullmsg="<%=word(125)%>"/>
	<td class="ykd"><div class="Validform_checktip"></div></td>
  </tr>
  <tr>
    <td style="border-right:none; border-bottom:none"></td>
    <td style="border-right:none;border-bottom:none" colspan="2"><input type="submit" value="<%=tw(289)%>" class="hyan"/>   </td>
  </tr>
</table>
</form></div>
<div id="con2">
<%

if hyzc=2 then
response.write(""&word(290)&"")
else
%>
<form action="<%=SitePath%>inc/MemberReg/?m=<%=m2%>&language=<%=lanyybs%>" method="post" class="checkform">
<table class="register">
<tr>
<td class="zkd" ><%=word(119)%></td>
<td class="ckd"><input name="MemName" type="text" id="MemName" ajaxurl="<%=sitepath%>ajax/yz.asp?language=<%=lanyybs%>"  class="inputxt" datatype="*6-16" nullmsg="<%=word(120)%>" errormsg="<%=word(121)%>"/><span class="fh">*</span></td>
<td class="ykd"><div class="Validform_checktip"><%=word(122)%></div></td>
</tr>

<tr>
<td class="zkd"><%=word(123)%></td>
<td class="ckd">
<input name="Password" type="password" id="Password"  class="inputxt" plugin="passwordStrength"  datatype="*6-18" errormsg="<%=word(124)%>" nullmsg="<%=word(125)%>"/><span class="fh">*</span>
<div class="passwordStrength"><b><%=word(126)%></b><span><%=word(127)%></span><span><%=word(128)%></span><span class="last"><%=word(129)%></span></div>
</td>
<td class="ykd"><div class="Validform_checktip"><%=word(130)%></div></td>
</tr>
<tr>
<td class="zkd"><%=word(131)%></td>
<td class="ckd"><input name="vPassword" type="password" id="vPassword"  class="inputxt" recheck="Password"  datatype="*6-18" errormsg="<%=word(132)%>" nullmsg="<%=word(133)%>"/><span class="fh">*</span></td>
<td class="ykd"><div class="Validform_checktip"><%=word(134)%></div></td>
</tr>
<tr>
<td class="zkd"><%=word(135)%></td>
<td class="ckd">

  <select name="Question" id="Question" class="inputxt" datatype="*" nullmsg="<%=word(136)%>">
                            <option value=""><%=word(137)%></option>
                        	<option value=1><%=word(138)%></option>
                        	<option value=2><%=word(139)%></option>
                        	<option value=3><%=word(140)%></option>
                        	<option value=4><%=word(141)%></option>
                        	<option value=5><%=word(142)%></option>
                        	<option value=6><%=word(143)%></option>
                        	<option value=7><%=word(144)%></option>
                        </select><span class="fh">*</span>
</td>
 <td class="ykd"><div class="Validform_checktip"><%=word(145)%></div></td>
</tr>
<tr>
<td class="zkd"><%=word(146)%></td>
<td class="ckd"><input name="Answer" type="text" id="Answer" class="inputxt" datatype="*" nullmsg="<%=word(147)%>"/><span class="fh">*</span>
</td>
<td class="ykd"><div class="Validform_checktip"><%=word(148)%></div></td>
</tr>
<tr>
<td class="zkd"><%=word(149)%></td>
<td class="ckd"><input name="RealName" type="text" id="RealName" class="inputxt" datatype="*2-6" nullmsg="<%=word(150)%>" errormsg="<%=word(151)%>" /><span class="fh">*</span></td>
<td class="ykd"><div class="Validform_checktip"><%=word(152)%></div></td>
</tr>
<tr>
<td class="zkd"><%=word(153)%></td>
<td class="ckd">
<select name="sex" id="sex" >
<option value="<%=word(154)%>"><%=word(154)%></option>
<option value="<%=word(155)%>"><%=word(155)%></option>
</select>
</td>
<td class="ykd"><div class="Validform_checktip"></div></td>
</tr>
<tr>
<td class="zkd"><%=word(156)%></td>
<td class="ckd"><input name="Company" type="text" id="Company"  class="inputxt"/></td>
<td class="ykd"><div class="Validform_checktip"></div></td>
</tr>
<tr>
<td class="zkd"><%=word(157)%></td>
<td class="ckd"><input name="Address" type="text" id="Address"  class="inputxt"/></td>
<td class="ykd"><div class="Validform_checktip"></div></td>
</tr>
<tr>
<td class="zkd"><%=word(158)%></td>
<td class="ckd"><input name="ZipCode" type="text" id="ZipCode"  class="inputxt"/></td>
<td class="ykd"><div class="Validform_checktip"></div></td>
</tr>
<tr>
<td class="zkd"><%=word(159)%></td>
<td class="ckd"><input name="Fax" type="text" id="Fax"  class="inputxt"/></td>
<td class="ykd"><div class="Validform_checktip"></div></td>
</tr>
<tr>
<td class="zkd"><%=word(160)%></td>
<td class="ckd"><input name="Mobile" type="text" id="Mobile"  class="inputxt" datatype="m" errormsg="<%=word(161)%>" nullmsg="<%=word(162)%>"/><span class="fh">*</span></td>
<td class="ykd"><div class="Validform_checktip"><%=word(163)%></div></td>
</tr>
<tr>
<td class="zkd"><%=word(164)%></td>
 <td class="ckd"><input name="Email" type="text" id="Email"  class="inputxt" ajaxurl="<%=sitepath%>ajax/yz.asp?language=<%=lanyybs%>"" datatype="e" errormsg="<%=word(165)%>" nullmsg="<%=word(166)%>"/><span class="fh">*</span></td>
 <td class="ykd"><div class="Validform_checktip"><%=word(167)%></div></td>
 </tr>

<tr>
<td class="zkd"><%=word(168)%></td>
<td class="ckd"><input name="captchacode" id="captchacode" type="text"  style="width:65px;" class="inputxt" datatype="*"  nullmsg="<%=word(169)%>" ajaxurl="<%=sitepath%>ajax/yzm.asp?language=<%=lanyybs%>""/> <img id="imgCaptcha" src="<%=sitepath%>inc/captcha.asp" />  <a href="javascript:; " onClick="RefreshImage('imgCaptcha')"><img src="<%=sitepath%>images/reload.png" width="20" height="20" /></a></td>
<td class="ykd"><div class="Validform_checktip"></div></td>
</tr>
<tr>
 <td  colspan="3"  style="border-right:none;border-bottom:none">
 <input name="qqapi" type="hidden" value="<%=Session("Openid")%>" />
 <input name="Submit" type="submit" value=" <%=word(170)%> " class="tjan"/>                   
 <input name="Reset" type="reset" value=" <%=word(171)%> " class="czan"/>                  
 </td>

</tr>

</table>
</form>
<%
end if

%>
</div>
</div>
</div>		
</div>
</div>

</div>
</div>

<!--#include file="../bottom.asp"-->

</body>
</html>

