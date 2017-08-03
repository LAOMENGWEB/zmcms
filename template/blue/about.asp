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




<div id="mbody">
<div class="mbody-in">
<div class="breadcrumb"><a href=/><%=tw(10)%></a><span class="sep">/</span><%=A("单页标题")%></div>
<div class="sidebar">

<dl class="l_pro">
<dt><%=lmmc(6)%></dt>
<%=info(1,1,"单页分类","zmone="&request.querystring("pone")&"",1,8,"","","",30,100,"yy","","id asc")%>
</dl>
</div>


  <div class="page-content">
        <h1><%=tw(298)%></h1>
        <div class="page-content-detail">
        <p>
        <%=A("单页内容")%>
        </p>
        </div>
    </div>






</div>
</div>

<!--#include file="bottom.asp"-->



</body>
</html>