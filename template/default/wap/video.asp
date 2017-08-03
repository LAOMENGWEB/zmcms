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
<%=tw(134)%>
</h3>
</div>
<%=infolist(4,2,"手机列表页视频","no","","","","",100,"d","",100,"id desc")%>
<div class="clear"></div>
</div>
<!--#include file="bottom.asp"--> 
</div>
<!--视频分类开始-->
<div class="sidebar_jsbox none"></div>
<div id="sidebar" class="none" data-csnow="25" data-class3="0">
<h3 class="title"><%=tw(135)%></h3>
<div id="wrapper_sidebar">
<div id="scroller_sidebar">
<%=lmdh("手机视频分类")%>
</div>
</div>
</div>
<!--图片分类结束-->
<script src="<%=sitepath%>wap/js/fun.inc.js" type="text/javascript"></script>
</body>
</html>