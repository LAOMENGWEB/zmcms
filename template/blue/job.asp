<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=tw(46)%>-<%=mc%>-<%=ym%></title>
<meta name="keywords" content="<%=gjc%>" />
<meta name="description" content="<%=ms%>" />
<link href="<%=template%>css/style.css" rel="stylesheet" type="text/css">
<!--#include file="include/js.asp"-->
</head>
<body ><ul>
	<li>
	<li>
</ul>
<!--#include file="top.asp"-->
<div id="mbody">
<div class="mbody-in">
<div class="breadcrumb"><a href=/><%=tw(10)%></a><span class="sep">/</span><%=tw(47)%></div>
<div class="sidebar">
<dl class="l_pro">
<dt><%=tw(48)%></dt>
<%=info(1,1,"单页分类","zmone="&zid&"",1,8,"","","",30,100,"yy","","id asc")%>
</dl>
</div>
<div class="cate-container">
<div class="list-container">      

<%=infolist(6,1,"列表页招聘","no","","","","",100,"yyyy/mm/dd","",100,"id desc")%>
</div>   
</div>
</div>
</div>



<!--#include file="bottom.asp"-->









</body>
</html>