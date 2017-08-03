<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=N("新闻标题")%>-<%=mc%>-<%=ym%></title>
<meta name="keywords" content="<%=N("新闻关键词")%>" />
<meta name="description" content="<%=N("新闻描述")%>" />
<link href="<%=template%>css/style.css" rel="stylesheet" type="text/css">
<!--#include file="include/js.asp"-->
</head>
<body >
<!--#include file="top.asp"-->

<div id="mbody">
<div class="mbody-in">
<div class="breadcrumb"><a href=/><%=tw(10)%></a><span class="sep">/</span><%=N("新闻标题")%></div>
<div class="sidebar">
<dl class="l_pro">
<dt><%=tw(11)%></dt>
<%=lmdh("新闻分类")%>
</dl>
</div>
   

 <div class="container">
        <h1 class="title"><%=N("新闻标题")%></h1>
        <div class="byline">
        
          <div class="mksx"> <%=tw(14)%>：<%=N("新闻作者")%>  　　<%=tw(15)%>：<%=N("发布时间")%>  　　<%=tw(16)%>：<%=N("新闻来源")%> 　　<%=tw(17)%>：<%=N("点击次数")%></div>
        </div>
        <div class="content">
            <p><%=N("新闻内容")%></p>
        </div>

        <div class="prev-next">
            <div class="prev"><%=tw(18)%>：<%=N("上条信息")%></div>
            <div class="next"><%=tw(19)%>：<%=N("下条信息")%></div>
        </div>
      
    </div>
 
</div>
</div>
</div>






</div>
</div>
<!--#include file="bottom.asp"-->
</body>
</html>