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
<div class="breadcrumb"><a href=/><%=tw(10)%></a><span class="sep">/</span><%=tw(21)%></div>
<div class="sidebar">
<dl class="l_pro">
<dt><%=tw(8)%></dt>
<%=lmdh("产品分类")%>
</dl>
</div>
<div class="cate-container">
<div class="list-container">  




<%=infolist(2,1,"列表页产品","no","","","","",100,"d","",100,"id desc")%>
</div>
</div>

</div>
</div>
<!--#include file="bottom.asp"-->
</body>
</html>

