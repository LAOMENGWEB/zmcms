<!-- #include file="../../Inc/sub.asp" --> 
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

<div class="sidebar">
<div class="sb_nav">
<h3 class='title'>
<p class="moresidebar"><%=tw(124)%></p>
<%=tw(128)%>
</h3>
</div>

<div id="newslist">
<ul>
<%=infolist(1,2,"手机列表页新闻","no","","","","",100,"yyyy-mm-dd","",100,"id desc")%>
</ul>
</div>

<div class="clear"></div>
</div>

<!--#include file="bottom.asp"--> 
</div>

<!--新闻分类开始-->
<div class="sidebar_jsbox none"></div>
<div id="sidebar" class="none" data-csnow="25" data-class3="0">
<h3 class="title"><%=tw(129)%></h3>
<div id="wrapper_sidebar">
<div id="scroller_sidebar">

<%=lmdh("手机新闻分类")%>

</div>
</div>
</div>
<!--新闻分类结束-->

<script src="<%=sitepath%>wap/js/fun.inc.js" type="text/javascript"></script>
</body>
</html>