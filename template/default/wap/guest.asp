<!DOCTYPE HTML>
<html>
<head>
<meta content='text/html; charset=utf-8' http-equiv='Content-Type' />
<meta charset='utf-8' />
<title><%=mc%> >> <%=ym%></title>
<meta content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=0" name="viewport" />
<meta content="yes" name="apple-mobile-web-app-capable" />
<meta content="black" name="apple-mobile-web-app-status-bar-style" />
<meta content="telephone=no" name="format-detection" />
<link rel="stylesheet" type="text/css" href="<%=template%>wap/css/default.css" />
<link rel="stylesheet" href="<%=template%>wap/awesome/css/font-awesome.min.css">
<link rel="stylesheet" type="text/css" href="<%=template%>wap/slider/slider.css" />
<link rel="stylesheet" type="text/css" href="<%=template%>wap/slider/slider.default.css" /> 
<!--#include file="../../../wap/include/js.asp"-->


</head>
<body>
<div class="zmcont">
<!--#include file="top.asp"--> 
<header>
<div class="top mtop">
<div class="mreturn"><a href="javascript:history.go(-1);"></a></div>
<div class="mtitle"><%=tw(136)%></div>
<div class="rightnav"></div>
</div>
</header>

<!--留言显示调用-->
<%=infolist(7,2,"手机列表页留言","no","","","","",100,"yyyy-mm-dd","",100,"id desc")%>

<!--发布留言调用-->       
<SCRIPT language=javaScript>
function Checkly()
{
if (document.form1.title.value==""){
layer.open({
    content: '<%=tw(137)%>',
    style: 'background-color:#09C1FF; color:#fff; border:none;',
    time: 2
});
document.form1.title.focus();
return false;
}
if (document.form1.phone.value==""){ 
layer.open({
    content: '<%=tw(138)%>',
    style: 'background-color:#09C1FF; color:#fff; border:none;',
    time: 2
});
document.form1.phone.focus(); 
return false; 
} 
if (document.form1.emaill.value==""){ 
layer.open({
    content: '<%=tw(139)%>',
    style: 'background-color:#09C1FF; color:#fff; border:none;',
    time: 2
});
document.form1.emaill.focus(); 
return false; 
} 	
if (document.form1.qq.value==""){ 
layer.open({
    content: '<%=tw(140)%>',
    style: 'background-color:#09C1FF; color:#fff; border:none;',
    time: 2
});
document.form1.qq.focus(); 
return false; 
} 
if (document.form1.content.value==""){ 
layer.open({
    content: '<%=tw(141)%>',
    style: 'background-color:#09C1FF; color:#fff; border:none;',
    time: 2
});
document.form1.content.focus(); 
return false; 
} 
}
</SCRIPT>
<div class="mobile_form ">
<form name="form1" method="post"  action="../../inc/mobilepost/?MemberID=0&m=<%=request.QueryString("m")%>&language=<%=language%>" onSubmit="return Checkly()">
<ul>
<li class="input required">
<div class="xbox"><input name='title' type='text' class='input-text' placeholder=<%=tw(142)%> /></div>
</li>
<li class="input">
<div class="xbox"><input name='phone' type='text' class='input-text' placeholder=<%=tw(143)%> /></div>
</li>
<li class="input">
<div class="xbox"><input name='emaill' type='text' class='input-text' placeholder=<%=tw(144)%> /></div>
</li>
<li class="input"><span class='info'><%=tw(145)%></span>
<div class="xbox"><input name='qq' type='text' class='input-text' placeholder=<%=tw(146)%> /></div>
</li>
<li class="input required">
<div class="xbox"><textarea name='content' class='textarea-text' placeholder=<%=tw(147)%>></textarea></div>
</li>
</ul>
<div class="input-submit"><input type='submit' name='Submit' value=<%=tw(148)%> /></div>
</form>
</div>
<!--#include file="bottom.asp"--> 
</div>
<script src="<%=sitepath%>wap/js/fun.inc.js" type="text/javascript"></script>
</body>
</html>