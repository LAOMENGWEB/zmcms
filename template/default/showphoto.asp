<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=C("图片标题")%><%=mc%>-<%=ym%></title>
<meta name="keywords" content="<%=C("图片关键词")%>" />
<meta name="description" content="<%=C("图片描述")%>" />
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
<dt><%=tw(29)%></dt>
<%=lmdh("图片分类")%>
</dl>
<dl class="l_pro">
<dt>日期归档</dt>
<%=guidang(3,20,1,"归档",3,0)%>

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
<div class="title"> <span class="fl"><%=tw(30)%></span>
<div class="fr"><%=tw(10)%>：<%=ddwz(3," >> ")%></div>
</div>
<div class="mainRightMain">


<div class="xwtitle"><%=C("图片标题")%><%=C("标题页数")%></div>
<div class="xwsx"> <%=tw(14)%>：<%=C("图片作者")%>  　　<%=tw(15)%>：<%=C("发布时间")%>  　　<%=tw(16)%>：<%=C("图片来源")%> 　　<%=tw(17)%>：<%=C("点击次数")%></div>

<div align="center" style="margin-top:20px">
<%=SHOWtpnr(2,2,7,65,500,400)%>
</div>

<div class="xwcontent">

<%=C("图片内容")%>

</div>
<div class="tybottom">
<div class="xwtags"><%=Ctags("标签：",1,"/")%></div>
<div class="xwnext">
<ul>
<li><%=tw(18)%>：<%=C("上条信息")%></li>
<li><%=tw(19)%>：<%=C("下条信息")%></li>
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