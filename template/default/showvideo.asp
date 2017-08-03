<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=T("视频标题")%><%=mc%>-<%=ym%></title>
<meta name="keywords" content="<%=tvkey%>" />
<meta name="description" content="<%=tvms%>" />
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
<dt><%=tw(31)%></dt>
<%=lmdh("视频分类")%>
</dl>
<dl class="l_pro">
<dt>日期归档</dt>
<%=guidang(4,20,1,"归档",2,0)%>

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
<div class="fyRight">
<div class="title"> <span class="fl"><%=tw(32)%></span>
<div class="fr"><%=tw(10)%>：<%=ddwz(4," >> ")%></div>
</div>
<div class="mainRightMain">


<div class="xwtitle"><%=T("视频标题")%><%=T("标题页数")%></div>
<div class="xwsx"> <%=tw(14)%>：<%=T("视频作者")%>  　　<%=tw(15)%>：<%=T("发布时间")%>  　　<%=tw(16)%>：<%=T("视频来源")%> 　　<%=tw(17)%>：<%=T("点击次数")%></div>

<div align="center" style="margin-top:20px; background:#000000;padding:20px 0">
<%=SHOWTV%>
</div>

<div class="xwcontent">

<%=T("视频内容")%>


</div>
<div class="tybottom">
<div class="xwtags"><%=Ttags("标签：",1,"/")%></div>
<div class="xwnext">
<ul>
<li><%=tw(18)%>：<%=T("上条信息")%></li>
<li><%=tw(19)%>：<%=T("下条信息")%></li>
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