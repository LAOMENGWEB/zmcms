<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=Z("下载标题")%>-<%=mc%>-<%=ym%></title>
<meta name="keywords" content="<%=Z("下载关键词")%>" />
<meta name="description" content="<%=Z("下载描述")%>" />
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
<dt><%=tw(33)%></dt>
<%=lmdh("下载分类")%>
</dl>
<dl class="l_pro">
<dt>日期归档</dt>
<%=guidang(5,20,1,"归档",2,0)%>

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
<div class="title"> <span class="fl"><%=tw(34)%></span>
<div class="fr"><%=tw(10)%>：<%=ddwz(5," >> ")%></div>
</div>
<div class="mainRightMain">



<div id="showdownsx">			
<ul class="paralist">               
<li><span class="name"><%=tw(35)%>：</span><%=Z("发布时间")%></li>                
<li> <span class="name"><%=tw(36)%>：</span>	<a href="<%=Z("下载地址一")%>">	<%=tw(39)%></a></li>
<li> <span class="name"><%=tw(37)%>：</span>	<a href="<%=Z("下载地址二")%>">	<%=tw(39)%></a></li>
<li> <span class="name"><%=tw(38)%>：</span>	<a href="<%=Z("下载地址三")%>">	<%=tw(39)%></a></li>
<li><span class="name"><%=tw(40)%>：</span><%=Z("适用系统")%></li> 
<li><span class="name"><%=tw(41)%>：</span><%=Z("软件语言")%></li>               
<li><span class="name"><%=tw(42)%>：</span><%=Z("软件类型")%></li>     
<li><span class="name"><%=tw(43)%>：</span><%=Z("软件大小")%></li>
<li><span class="name"><%=tw(44)%>：</span><%=Z("点击次数")%></li>	

</ul>			


</div>	





<div class="showdownsmbh">
<div class="showdownsm">
<h3 class="downtitle"><span><%=tw(45)%></span></h3>
<%=Z("下载内容")%>
</div>
<div class="downtags"><%=Ztags("标签：",1,"/")%></div><map name="">
	<area shape="" href="" coords="" alt="">
</map>
<div class="downnext">
<ul>
<li><%=tw(18)%>：<%=Z("上条信息")%></li>
<li><%=tw(19)%>：<%=Z("下条信息")%></li>
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