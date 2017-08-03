<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="keywords" content="<%=gjc%>" />
<meta name="description" content="<%=ms%>" />
<title><%=mc%>-<%=ym%></title>
<link href="<%=template%>css/style.css" rel="stylesheet" type="text/css">
<!--#include file="include/js.asp"-->
<script src="<%=template%>img/jquery.carouFredSel.js"></script>
</head>
<body>
<!--#include file="top.asp"-->
<!--PRODUCT S-->

<div class="default-products-title">
<h3>PRODUCTS EXHIBITION</h3>
<h2>产品展示</h2>
<div class="main" style=" margin-top:50px">
<%=info(3,1,"首页产品中心","no","",30,"","","",30,100,"yyyy-mm-dd","","id desc")%>
<div class="default-products-more">
<%if yy=1 then%>
<a href="<%=product%>/?m=3" target="_blank">
<%else%>
<a href="<%=product%>/<%=ProductClass%>.html" target="_blank">
<%end if%>
<%=tw(1)%></a></div>
</div>

<script type="text/javascript">
 $(window).load(function() {
  $("#scrollable-29326").carouFredSel({
   infinite  : false,
   circular  : false,
   auto   : false,
   width    : '100%',

   prev   : {button:"#carousel-prev-29326"},
   next   : {button:"#carousel-next-29326"}
  });
 });
</script>
</div>

<!--PRODUCT E-->

<!--NEWS S-->
<div class="default-news-title">
<h3>LATEST NEWS</h3>
<h2>最新资讯</h2>
<div class="main" style=" margin-top:50px">
<%=info(2,1,"首页新闻中心","no","",4,"","","",25,100,"yyyy-mm-dd","","id desc")%>

<div class="default-news-more">
<%if yy=1 then%>
<a href="<%=iinfo%>/?m=2" target="_blank">
<%else%>
<a href="<%=iinfo%>/<%=Newclass%>.html" target="_blank">
<%end if%>
<%=tw(1)%></a></div>
</div>
</div>
<!--NEWS E-->
<!--ABOUT S-->
<div class="default-abouts-title">
<h3>ABOUT OURS</h3>
<h2>关于我们</h2>
<div class="main" style=" margin-top:50px">

<%=info(1,1,"首页企业简介","id="&qyid&"",1,10,"","","",30,660,"yy","","id desc")%>


</div></div>
<!--ABOUT E-->





<!--CASE S-->

<div class="default-case-title">
<h3>CASE EXHIBITION</h3>
<h2>案例展示</h2>

<div class="main" style=" margin-top:50px">
<%=info(4,1,"首页图片中心","no","",20,"","","",30,100,"yyyy-mm-dd","","id desc")%>
<div class="default-case-more">

<%if yy=1 then%>
<a href="<%=photo%>/?m=4" target="_blank">
<%else%>
<a href="<%=photo%>/<%=alClass%>.html" target="_blank">
<%end if%>

<%=tw(1)%></a></div>
</div>
</div>

<!--CASE E-->


<!--Content S-->

<div class="default-content">
<div class="main">
<div class="lxwm">
<div class="title">
<span><%=tw(3)%></span>
</div>    

<ul>
<li><%=tw(4)%>：<%=lxdz%></li>
<li><%=tw(5)%>：<%=sj%></li>
<li><%=tw(6)%>：<%=cz%></li>
<li><%=tw(7)%>：<%=lxr%></li>
</ul>
</div>
<div class="deault-cpzs">
<div class="title">
<span><%=tw(8)%></span>
</div>    
<ul>
<%=lmdh("首页产品分类")%>
</ul>
</div>

<div class="ewm">
<div class="title">
<span><%=tw(9)%></span>
</div>    

<ul>
<li><img src="<%=ewm%>" width="<%=ewmwidth%>" height="<%=ewmheight%>"></li>

</ul>
</div>
</div>
</div>
<!--content E-->

<!--#include file="bottom.asp"-->

</body>
</html>