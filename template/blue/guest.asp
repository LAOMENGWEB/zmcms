<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=mc%>-<%=ym%></title>
<meta name="keywords" content="<%=gjc%>" />
<meta name="description" content="<%=ms%>" />
<link href="<%=template%>css/style.css" rel="stylesheet" type="text/css">

<!--#include file="include/js.asp"-->


</head>
<body >
<!--#include file="top.asp"-->
<div id="mbody">
<div class="mbody-in">
<div class="breadcrumb"><a href=/><%=tw(10)%></a><span class="sep">/</span><%=tw(62)%></div>
<div class="sidebar">
<dl class="l_pro">
<dt><%=tw(63)%></dt>

<dd>
<a href="<%=sitepath%>guest/?m=9&language=<%=language%>" title="<%=tw(64)%>"  <% if request.QueryString("m")=9 then response.write("class=""select""") end if%>><%=tw(64)%></a>
</dd>

<dd><a href="<%=sitepath%>showguest/?m=9&language=<%=language%>" title="<%=tw(65)%>" ><%=tw(65)%></a></dd>





</dl>
</div>
   
<!--右侧-->
<div class="container">
<!--发布留言代码开始-->
<form name="form1" method="post"  action="<%=sitepath%>inc/post/?MemberID=<%=mMemID%>&m=<%=request.QueryString("m")%>&language=<%=language%>" class="checkform">

<table  class="Fguest" cellpadding="0" cellspacing="0">
<tr>
<td class="zkd"><%=word(72)%></td>
<td class="ckd"><input name="title"  type="text" value="<%=mRealName%>" class="inputxt" datatype="*" nullmsg="<%=word(73)%>" /><span class="fh">*</span></td>
<td class="ykd"><div class="Validform_checktip"><%=word(74)%></div></td>
</tr>
<tr>
<td  class="zkd"><%=word(75)%></td>
<td  class="ckd"><input name="phone" type="text" value="<%=mMobile%>" class="inputxt" datatype="m" nullmsg="<%=word(76)%>" errormsg="<%=word(77)%>"/><span class="fh">*</span></td>
<td  class="ykd"><div class="Validform_checktip"><%=word(78)%></div></td>
</tr>
<tr>
<td class="zkd"><%=word(79)%></td>
<td class="ckd"><input name="emaill" type="text" value="<%=mEmail%>" class="inputxt" datatype="e" nullmsg="<%=word(80)%>" errormsg="<%=word(81)%>"/><span class="fh">*</span></td>
<td class="ykd"><div class="Validform_checktip"><%=word(82)%></div></td>
</tr>
<tr>
<td class="zkd"><%=word(83)%></td>
<td class="ckd"><input name="qq" type="text" class="inputxt" datatype="n" nullmsg="<%=word(84)%>" errormsg="<%=word(85)%>"/><span class="fh">*</span></td>
<td class="ykd"><div class="Validform_checktip"><%=word(86)%></div></td>
</tr>
<tr>
<td class="zkd"><%=word(87)%></td>
<td class="ckd"><textarea class="textarea" name="content" cols="60" rows="8" id="content"   datatype="*" nullmsg="<%=word(88)%>" ajaxurl="<%=sitepath%>ajax/bad.asp?&language=<%=language%>"></textarea><span class="fh">*</span></td>
<td class="ykd"><div class="Validform_checktip"><%=word(89)%></div></td>
</tr>
<%if qq=1 then%>
<tr>
<td class="zkd"><%=word(90)%></td>
<td class="ckd"><%=qqh%></td>
<td class="ykd"><div class="Validform_checktip"></div></td>
</tr>
<%end if%>
<tr>
<td class="zkd" ><%=word(91)%></td>
<td class="ckd"><input  name="captchacode" type="text" id="captchacode" size="6" maxlength="6" style="width:65px;" class="inputxt" datatype="*"  nullmsg="<%=word(92)%>" ajaxurl="<%=sitepath%>ajax/yzm.asp?&language=<%=language%>"/> <img id="imgCaptcha" src="<%=SITEPATH%>inc/captcha.asp" /> <a href="###" onClick="RefreshImage('imgCaptcha')"><img src="<%=sitepath%>images/reload.png" width="20" height="20" /></a></td>
<td class="ykd"><div class="Validform_checktip"></div></td>
</tr>
<tr>
<td colspan="3"  class="last">
<input name="Submit" type="submit" class="tjan" value="<%=word(93)%> " />
<input name="Reset" type="reset" class="czan" value=" <%=word(94)%> " /></td>
</tr>
 </table>
</form>





<!--发布留言代码结束-->

</div>

</div>
</div>

<!--#include file="bottom.asp"-->



</body>
</html>
