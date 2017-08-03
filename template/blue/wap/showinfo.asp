<!DOCTYPE HTML>
<html>
<head>
<meta content='text/html; charset=utf-8' http-equiv='Content-Type' />
<meta charset='utf-8' />
<title><%=N("新闻标题")%>-<%=mc%>-<%=ym%></title>
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
<div class="mtitle"><%=tw(160)%></div>
<div class="rightnav"></div>
</div>
</header>

<!--内容区-->
<div id="shownews">
<h1 class="title"><%=N("新闻标题")%><p class="hits"><%=N("发布时间")%><%=tw(161)%>：<%=N("点击次数")%></p></h1>
<div class="editor" id="Gallery1">

<%=N("wap新闻内容")%>

</div>
<div class="clear"></div>
</div>

<!--上下条-->
<div class="zm_page">
<ul class='preul'>
<li class='preinfo'>
<p class='pret'><%=tw(158)%></p>
<p class='pres'><%=N("WAP上条信息")%></p>
</li>
<li class='nextinfo'>
<p class='pret'><%=tw(159)%></p>
<p class='pres'><%=N("WAP下条信息")%></p>
</li>
</ul>
<div class="clear"></div>
</div>	
<!--#include file="bottom.asp"--> 
</div>

<script src="<%=sitepath%>wap/js/fun.inc.js" type="text/javascript"></script>
</body>
</html>