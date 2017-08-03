<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="keywords" content="<%=gjc%>" />
<meta name="description" content="<%=ms%>" />
<title><%=mc%>-<%=ym%></title>
<link href="<%=template%>css/style.css" rel="stylesheet" type="text/css">
<!--#include file="include/js.asp"-->

</head>
<body>
<!--#include file="top.asp"-->
<!--主内容区-->
<div id="mbody">
<div class="wrapper">
<!--新闻中心-->
<div class="index-news" id="index_news">
<h3 class="header">

<%if yy=1 then%>
<a href="<%=iinfo%>/?m=2" target="_blank" class="more">
<%else%>
<a href="<%=iinfo%>/<%=Newclass%>.html" target="_blank" class="more">
<%end if%>
MORE</a>



<%=tw(12)%></h3>
<div class="section">
<ul>
<%=info(2,1,"首页新闻中心","no","",4,"","","",50,100,"yyyy-mm-dd","","id desc")%>
</ul>
</div>
</div>
<!--产品中心-->
<div class="index-service" id="index_service">
<h3 class="header">
<%if yy=1 then%>

<a href="<%=product%>/?m=3" target="_blank" class="more">
<%else%>
<a href="<%=product%>/<%=ProductClass%>.html" target="_blank" class="more">
<%end if%>
MORE</a>

<%=tw(21)%></h3>
<div class="section">
<ul>
<%=info(3,1,"首页产品中心","no","",6,"","","",30,100,"yyyy-mm-dd","","id desc")%>
</ul>
</div>
</div>
<!--企业简介-->
<div class="index-about">
<h3><a href="?about_us.html" ><%=tw(48)%></a></h3>
<%=info(1,1,"首页企业简介","id="&qyid&"",1,10,"","","",30,450,"yy","","id desc")%>
    </div>

    <div class="clear"></div>
</div>
</div>
<!--#include file="bottom.asp"-->
</body>
</html>