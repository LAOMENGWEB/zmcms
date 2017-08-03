<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=A("单页标题")%></title>
<meta name="keywords" content="<%=A("单页关键字")%>" />
<meta name="description" content="<%=A("单页描述")%>" />
<link href="<%=template%>css/style.css" rel="stylesheet" type="text/css">
<!--#include file="include/js.asp"-->
</head>
<body >
<!--#include file="top.asp"-->

<div class="wrap">
<div class="main">
 <!--左侧-->
<div class="fyLeft">

<dl class="l_pro">
<dt><%=lmmc(6)%></dt>
<%=info(1,1,"单页分类","zmone="&request.querystring("pone")&"",1,8,"","","",30,100,"yy","","id asc")%>
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
<div class="title"> <span class="fl"><%=A("单页标题")%></span>
<div class="fr"><%=tw(10)%>：<%=A("单页标题")%></div>
</div>
<div class="mainRightMain">
<div class="xwcontent">
<%=A("单页内容")%>
</div>
</div>

</div>
</div>
</div>


<div>
</div>
<!--#include file="bottom.asp"-->



</body>
</html>