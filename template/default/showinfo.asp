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
<div class="wrap">
<div class="main">
 <!--左侧-->
<div class="fyLeft">
<dl class="l_pro">
<dt><%=tw(11)%></dt>
<%=lmdh("新闻分类")%>
</dl>
<dl class="l_pro">
<dt>日期归档</dt>
<%=guidang(1,20,1,"归档",2,0)%>

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
<div class="title"> <span class="fl"><%=tw(12)%></span>
<div class="fr"><%=tw(10)%>：<%=ddwz(1," >> ")%></div>
</div>
<div class="mainRightMain">


<div class="xwtitle"><%=N("新闻标题")%><%=N("标题页数")%></div>
<div class="xwsx"> <%=tw(14)%>：<%=N("新闻作者")%>  　　<%=tw(15)%>：<%=N("发布时间")%>  　　<%=tw(16)%>：<%=N("新闻来源")%> 　　<%=tw(17)%>：<%=N("点击次数")%></div>
<div class="xwcontent">

<%=N("新闻内容")%>


</div>
<div class="tybottom">
<div class="xwtags"><%=Ntags("标签：",1,"/")%></div>
<div class="xwnext">
<ul>
<li><%=tw(18)%>：<%=N("上条信息")%></li>
<li><%=tw(19)%>：<%=N("下条信息")%></li>
</ul>
</div>


</div>
</div>
</div>
</div>
</div>
<!--#include file="bottom.asp"-->
</body>
</html>