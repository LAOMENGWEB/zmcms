<!DOCTYPE HTML>
<html>
<head>
<meta content='text/html; charset=utf-8' http-equiv='Content-Type' />
<meta charset='utf-8' />
<title><%=mc%></title>
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
<!--公司简介开始代码-->
<div class="style-1">
<h3 class='title '><%if language="ch" then%><%if sqlno=1 then%><a href="<%=sitepath%>wap/about/?pone=1&aboutID=1&language=<%=language%>&m=1"><%else%><a href="<%=sitepath%>wap/about/?pone=1&aboutID=3&language=<%=language%>&m=1"><%end if%><%else%><%if sqlno=1 then%><a href="<%=sitepath%>wap/about/?pone=2&aboutID=2&language=<%=language%>"><%else%><a href="<%=sitepath%>wap/about/?pone=2&aboutID=4&language=<%=language%>"><%end if%><%end if%><%=tw(121)%><span class="more"></span></a></h3>
<div class="editor">

<%=info(1,2,"首页手机公司简介","id="&qyid&"",1,10,"","","",30,200,"yy","","id desc")%>
<div class="clear"></div>
</div>
</div>
<!--公司简介结束代码-->


<!--产品展示开始-->
<div class=" style-1">
<h3 class='title '><a href="<%=sitepath%>wap/product/?m=3&language=<%=language%>"><%=tw(122)%><span class="more"></span></a></h3>
<div class="dl-jqrun">
<%=info(3,2,"首页手机产品中心","no","",10,"","","",30,100,"yy","","id desc")%>
<div class="clear"></div>
</div>
</div>
<!--产品展示结束-->

<!--新闻开始-->
<div class="newslist style-1">
<h3 class='title '><a href="<%=sitepath%>wap/info/?m=2&language=<%=language%>"><%=tw(123)%><span class="more"></span></a></h3>
<ul>
<%=info(2,2,"首页手机新闻中心","no","",8,"","","",30,100,"yyyy/mm/dd","","id desc")%>
</ul>
</div>
<!--新闻结束-->

<!--#include file="bottom.asp"--> 
</div>
<script src="<%=sitepath%>wap/js/fun.inc.js" type="text/javascript"></script>
</body>
</html>