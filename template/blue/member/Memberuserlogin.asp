<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=mc%>-<%=tw(70)%>-<%=tw(109)%></title>
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
<div class="fr"><%=tw(10)%>：<%=tw(70)%> >> <%=tw(110)%></div>
</div>


<div class="mainRightMain">


 <form name="formLogin" method="post" action="<%=sitepath%>inc/MemberLogin/?m=<%=request.QueryString("m")%>&language=<%=language%>" class="registerform">
  <table  class="hydl">
  <tr>
    <td class="zkd" align="right"><%=tw(111)%>：</td>
    <td style="border-right:none"><input type="text" name="LoginName" id="LoginName" class="wbk"/> <a href="<%=sitepath%>member/MemberRegister/?m=<%=request.QueryString("m")%>"><%=tw(112)%></a></td>
  </tr>
  <tr>
    <td align="right"><%=tw(113)%>：</td>
    <td style="border-right:none"><input type="password" name="LoginPassword" id="LoginPassword" class="wbk"/> <a href="<%=sitepath%>member/MemberGetPass/?m=<%=request.QueryString("m")%>&language=<%=language%>"><%=tw(114)%></a></td>
  </tr>
  <tr>
    <td align="right" ><%=tw(115)%>：</td>
    <td style="border-right:none">
      <input name="captchacode" type="text" class="wbk" id="captchacode" style="width:98px;" size="6" maxlength="6"/> <img id="imgCaptcha" src="<%=sitepath%>inc/captcha.asp" /> <a href="###" onClick="RefreshImage('imgCaptcha')"><img src="<%=sitepath%>images/reload.png" width="20" height="20" /></a>
       </td>
	 </tr>
  <tr>
    <td style="border-right:none; border-bottom:none"></td>
    <td style="border-right:none;border-bottom:none"><input type="submit" value="<%=tw(116)%>" class="hyan"/> </td>
  </tr>
  
    <tr>
    <td style="border-right:none; border-bottom:none"></td>
    <td style="border-right:none;border-bottom:none">   <a href="<%=sitepath%>plug/qqapi"><img src="../../images/qq_login.gif" /></a></td>
  </tr>
</table>
</form>
  
</div>
</div>

</div>
</div>

<!--#include file="../bottom.asp"-->

 


</body>
</html>