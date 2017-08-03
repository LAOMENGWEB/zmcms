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
<div id="mbody">
<div class="mbody-in">
<div class="breadcrumb"><a href=/><%=tw(10)%></a><span class="sep">/</span><%=C("图片标题")%></div>
<div class="sidebar">
<dl class="l_pro">
<dt><%=tw(29)%></dt>
<%=lmdh("图片分类")%>
</dl>
</div>


    <div class="container">
        <h1 class="title"><%=C("图片标题")%></h1>
        <div class="byline">
       <div class="mksx"> <%=tw(14)%>：<%=C("图片作者")%>  　　<%=tw(15)%>：<%=C("发布时间")%>  　　<%=tw(16)%>：<%=C("图片来源")%> 　　<%=tw(17)%>：<%=C("点击次数")%></div>
        </div>
       
        
        <div align=center class="contentimg">
      <%=SHOWtpnr(2,2,7,65,500,400)%>
        </div>
   
        <div class="content">
        <%=C("图片内容")%>
        </div>

      
        <div class="prev-next">
            <div class="prev"><%=tw(18)%>：<%=C("上条信息")%></div>
            <div class="next"><%=tw(19)%>：<%=C("下条信息")%></div>
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