<!DOCTYPE HTML>
<html>
<head>
<meta content='text/html; charset=utf-8' http-equiv='Content-Type' />
<meta charset='utf-8' />
<title><%=Z("下载标题")%>-<%=mc%>-<%=ym%></title>
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
<div class="mtitle"><%=tw(149)%></div>
<div class="rightnav"></div>
</div>
</header>



<div id="showjob">
<div class="paralist">
<ul>
<li><%=tw(150)%>：<%=Z("发布时间")%></li>                

<li><%=tw(151)%>：<%=Z("适用系统")%></li> 
<li><%=tw(152)%>：<%=Z("软件语言")%></li>               
<li><%=tw(153)%>：<%=Z("软件类型")%></li>     
<li><%=tw(154)%>：<%=Z("软件大小")%></li>
<li><%=tw(155)%>：<%=Z("点击次数")%></li>	
</ul>
</div>
</div>

<div class="info_cv">
<div class="button">
<a href="<%=Z("下载地址一")%>"><%=tw(156)%></a>
</div>
</div>

<div class="view"><%=tw(157)%></div>
<!--下载详情-->
<div class="editor" id="Gallery1">

<%=Z("wap下载内容")%>
</div>

<!--上下条-->
<div class="zm_page">
<ul class='preul'>
<li class='preinfo'>
<p class='pret'><%=tw(158)%></p>
<p class='pres'><%=Z("WAP上条信息")%></p>
</li>
<li class='nextinfo'>
<p class='pret'><%=tw(159)%></p>
<p class='pres'><%=Z("WAP下条信息")%></p>
</li>
</ul>
<div class="clear"></div>
</div>	

<!--#include file="bottom.asp"--> 
</div>
<script src="<%=sitepath%>wap/js/fun.inc.js" type="text/javascript"></script>
</body>
</html>