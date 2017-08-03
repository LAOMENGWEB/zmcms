
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=mc%>-<%=word(70)%>-<%=word(107)%></title>
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

</div>

<!--右侧-->
<div class="fyRight">
<div class="title"> <span class="fl"><%=tw(70)%></span>
<div class="fr"><%=tw(10)%>：<%=tw(70)%> >> <%=tw(108)%></div>
</div>


<div class="mainRightMain">

<!--会员注册代码开始-->
<%
if hyzc=2 then
writemsg1 2,""&word(118)&"","","wrong"
else
%>
<form action="<%=SitePath%>inc/MemberReg/?m=<%=request.QueryString("m")%>&language=<%=language%>" method="post" class="checkform">
<table class="register">
<tr>
<td class="zkd" ><%=word(119)%></td>
<td class="ckd"><input name="MemName" type="text" id="MemName" ajaxurl="<%=sitepath%>ajax/yz.asp?language=<%=language%>"  class="inputxt" datatype="*6-16" nullmsg="<%=word(120)%>" errormsg="<%=word(121)%>"/><span class="fh">*</span></td>
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
 <td class="ckd"><input name="Email" type="text" id="Email"  class="inputxt" ajaxurl="<%=sitepath%>ajax/yz.asp?language=<%=language%>" datatype="e" errormsg="<%=word(165)%>" nullmsg="<%=word(166)%>"/><span class="fh">*</span></td>
 <td class="ykd"><div class="Validform_checktip"><%=word(167)%></div></td>
 </tr>

<tr>
<td class="zkd"><%=word(168)%></td>
<td class="ckd"><input name="captchacode" id="captchacode" type="text"  style="width:65px;" class="inputxt" datatype="*"  nullmsg="<%=word(169)%>" ajaxurl="<%=sitepath%>ajax/yzm.asp?language=<%=language%>"/> <img id="imgCaptcha" src="<%=sitepath%>inc/captcha.asp" />  <a href="###" onClick="RefreshImage('imgCaptcha')"><img src="<%=sitepath%>images/reload.png" width="20" height="20" /></a></td>
<td class="ykd"><div class="Validform_checktip"></div></td>
</tr>
<tr>
 <td  colspan="3"  style="border-right:none;border-bottom:none">
 <input name="Submit" type="submit" value=" <%=word(170)%> " class="tjan"/>                   
 <input name="Reset" type="reset" value=" <%=word(171)%> " class="czan"/>                  
 </td>

</tr>

</table>
</form>
<%
end if
%>

<!--会员注册代码结束-->






</div>
</div>

</div>
</div>




<!--#include file="../bottom.asp"-->
</body>
</html>

