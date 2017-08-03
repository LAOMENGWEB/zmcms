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
<div class="breadcrumb"><a href=/><%=tw(10)%></a><span class="sep">/</span><%=tw(61)%></div>
<div class="sidebar">
<dl class="l_pro">
<dt><%=tw(48)%></dt>
<%=info(1,1,"单页分类","zmone="&zid&"",1,8,"","","",30,100,"yy","","id asc")%>
</dl>
</div>
   
<!--右侧-->
<div class="container">

<!--提交简历开始-->
<%
if session("MemName")="" and session("MemLogin")<>"Succeed" and zxyp=2 then
writemsg1 2,""&word(301)&"","member/memberuserlogin/?m="&m2&"&language="&language&"","wrong"
else
Quarters = request.QueryString("Quarters")
%>
<form  method="post"  action="<%=sitepath%>inc/tjjl/?MemberID=<%=mMemID%>&action=add&m=<%=request.QueryString("m")%>&language=<%=language%>" class="checkform">
<table class="tjjl">
<tr>
<td class="zkd"><%=word(302)%></td>
<td class="ckd"><input name="Quarters" type="text"  id="OrderName" value="<%=Quarters%>" class="inputxt"/><span class="fh">*</span></td>
<td class="ykd"><div class="Validform_checktip"><%=word(303)%></div></td>
</tr>
<tr >
<td class="zkd"> <%=word(304)%></td>
<td class="ckd"><input  name="Name" value="<%=mrealname%>" class="inputxt" datatype="*2-6" nullmsg="<%=word(305)%>" errormsg="<%=word(306)%>"/><span class="fh">*</span></td>
<td class="ykd"><div class="Validform_checktip"><%=word(307)%></div></td>
</tr>
<tr>
<td class="zkd" ><%=word(308)%></td>
<td class="ckd">
<input name="Sex"   type="radio"  value="<%=word(309)%>" <%if mSex=""&word(309)&"" then response.write ("checked")%>  style="border:none"/> <%=word(309)%>
<input type="radio" name="Sex"    value="<%=word(310)%>"  <%if mSex=""&word(310)&"" then response.write ("checked")%> style="border:none"/> <%=word(310)%>
</td>
<td class="ykd"><div class="Validform_checktip"></div></td>
</tr>
<tr>
<td   class="zkd" ><%=word(311)%></td>
<td class="ckd">
<select id="Marry"   name="Marry" >
<option value="<%=word(312)%>"  selected><%=word(312)%></option>
<option value="<%=word(313)%>"><%=word(313)%></option>
</select>
</td>
<td class="ykd"><div class="Validform_checktip"></div></td>
</tr>
<tr>
<td   class="zkd"><%=word(314)%></td>
<td class="ckd"><input name="Birthday" class="inputxt" datatype="*" nullmsg="<%=word(315)%>"><span class="fh">*</span></td>
<td class="ykd"><div class="Validform_checktip"><%=word(316)%></div></td>
</tr>
<tr>
<td    class="zkd"><%=word(317)%></td>
<td class="ckd"><input name="Stature" id="stature"  class="inputxt" datatype="n" nullmsg="<%=word(318)%>" errormsg="<%=word(319)%>"/><span class="fh">*</span></td>
<td class="ykd"><div class="Validform_checktip"><%=word(320)%></div></td>
</tr>
<tr>
<td class="zkd" ><%=word(321)%></td>
<td class="ckd"><input name="School"  class="inputxt" datatype="*" nullmsg="<%=word(322)%>"/><span class="fh">*</span></td>
<td class="ykd"><div class="Validform_checktip"><%=word(323)%></div></td>
</tr>
<tr>
<td class="zkd" ><%=word(324)%></td>
<td class="ckd"><input name="Studydegree" class="inputxt" datatype="*" nullmsg="<%=word(325)%>"/><span class="fh">*</span></td>
<td class="ykd"><div class="Validform_checktip"><%=word(326)%></div></td>
</tr>
<tr>
<td  class="zkd" ><%=word(327)%></td>
<td class="ckd"><input name="Specialty" class="inputxt" datatype="*" nullmsg="<%=word(328)%>"/><span class="fh">*</span></td>
<td class="ykd"><div class="Validform_checktip"><%=word(329)%></div></td>
</tr>
<tr>
<td    class="zkd"><%=word(330)%></td>
<td class="ckd"><input name="Gradyear" class="inputxt" datatype="*" nullmsg="<%=word(331)%>"/><span class="fh">*</span></td>
<td class="ykd"><div class="Validform_checktip"><%=word(332)%></div></td>
</tr>
<tr>
<td  class="zkd"><%=word(333)%></td>
<td class="ckd"><input name=Residence id=Residence class="inputxt" datatype="*" nullmsg="<%=word(334)%>"/><span class="fh">*</span></td>
<td class="ykd"><div class="Validform_checktip"><%=word(335)%></div></td>
</tr>
 <tr>
<td    class="zkd"><%=word(336)%></td>
<td class="ckd">
<textarea id=Edulevel class="textarea" name=Edulevel datatype="*" nullmsg="<%=word(337)%>"></textarea><span class="fh">*</span></td>
<td class="ykd"><div class="Validform_checktip"><%=word(338)%></div></td>
</td>
 </tr>
  <tr>
<td  class="zkd" ><%=word(339)%></td>
<td class="ckd">
<textarea  id=Experience class="textarea" name=Experience datatype="*" nullmsg="<%=word(340)%>"></textarea><span class="fh">*</span> </td>
<td class="ykd"><div class="Validform_checktip"><%=word(341)%></div></td>
</td>
</tr>
<tr >
<td   class="zkd" ><%=word(342)%></td>
<td class="ckd"><input name=Phone id=Phone  value="<%=mMobile%>" class="inputxt" datatype="m" errormsg="<%=word(343)%>" nullmsg="<%=word(344)%>"/><span class="fh">*</span></td>
<td class="ykd"><div class="Validform_checktip"><%=word(345)%></div></td>
</tr>
<tr>
<td  class="zkd" ><%=word(346)%></td>
<td class="ckd"><input  name=Email id=Email value="<%=mEmail%>" class="inputxt" datatype="e" errormsg="<%=word(347)%>" nullmsg="<%=word(348)%>"/><span class="fh">*</span></td>
<td class="ykd"><div class="Validform_checktip"><%=word(349)%></div></td>
</tr>
<tr >
<td   class="zkd"><%=word(350)%></td>
<td class="ckd"><input name=Add id=Add value="<%=mAddress%>" class="inputxt" datatype="*" nullmsg="<%=word(351)%>"/><span class="fh">*</span></td>
<td class="ykd"><div class="Validform_checktip"><%=word(352)%></div></td>
</tr>
<tr>
<td  class="zkd" ><%=word(353)%></td>
<td class="ckd">
<input name=Postcode id=Postcode value="<%=mZipCode%>" class="inputxt" datatype="p" nullmsg="<%=word(354)%>" errormsg="<%=word(355)%>" /><span class="fh">*</span></td>
<td class="ykd"><div class="Validform_checktip"><%=word(356)%></div></td>
</tr>
<tr>
<td class="zkd"><%=word(357)%></td>
<td class="ckd"><input class="inputxt" type="text" style="width:65px;" name="captchacode"  id="captchacode" datatype="*"  nullmsg="<%=word(358)%>"  ajaxurl="<%=sitepath%>ajax/yzm.asp?language=<%=language%>"/> <img id="imgCaptcha" src="<%=sitepath%>inc/captcha.asp" /> <a href="###" onClick="RefreshImage('imgCaptcha')"><img src="<%=sitepath%>images/reload.png" width="20" height="20" /></a>           </td>
<td class="ykd"><div class="Validform_checktip"></div></td>
</tr>
<tr >
<td colspan="3" class="last">
<input class="tjan" type="submit" name="submit8" value=" <%=word(359)%> " />
<input class="czan" type="reset" name="Submit8" value=" <%=word(360)%> " /></td>
</tr>
</table>

</form>
<%end if%>

<!--提交简历结束-->

</div>

</div>
</div>
<!--#include file="bottom.asp"-->
</body>
</html>
