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

<div id="mbody">
<div class="mbody-in">
<div class="breadcrumb"><a href=/><%=tw(10)%></a><span class="sep">/</span><%=Z("下载标题")%></div>
<div class="sidebar">
<dl class="l_pro">
<dt><%=tw(33)%></dt>
<%=lmdh("下载分类")%>
</dl>
</div>
<!--右侧-->







 <div class="container">
        <h1 class="title"><%=Z("下载标题")%></h1>
         <div class="contentimg">
      
<div >			
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

        </div>
        <div class="content">
            <p><%=Z("下载内容")%></p>
        </div>

        <div class="prev-next">
            <div class="prev"><%=tw(18)%>：<%=Z("上条信息")%></div>
            <div class="next"><%=tw(19)%>：<%=Z("下条信息")%></div>
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