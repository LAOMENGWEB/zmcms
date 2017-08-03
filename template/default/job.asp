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
<div class="wrap">
<div class="main">
 <!--左侧-->
<div class="fyLeft">
<dl class="l_pro">
<dt><%=tw(48)%></dt>
<%=info(1,1,"单页分类","zmone="&zid&"",1,8,"","","",30,100,"yy","","id asc")%>
</dl>
<dl class="l_content">
<dt><%=tw(4)%>：<%=lxdz%></dt>
</dl>
<dl class="l_content">
<dt><%=tw(5)%>：<%=sj%></dt>
</dl>
<dl class="l_content">
<dt><%=tw(6)%>：<%=cz%></dt>
</dl>
<dl class="l_content">
<dt><%=tw(7)%>：<%=lxr%></dt>
</dl>

</div>
   
<!--右侧-->
<div class="fyRight">
<div class="title"> <span class="fl"><%=tw(47)%></span>
<div class="fr"><%=tw(10)%>：<%=ddwz(6," >> ")%></div>
</div>


<div class="mainRightMain">

<%=infolist(6,1,"列表页招聘","no","","","","",100,"yyyy/mm/dd","",100,"id desc")%>
</div>
</div>

</div>
</div>

<!--#include file="bottom.asp"-->
</body>
</html>