<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=P("产品标题")%><%=mc%>-<%=ym%></title>
<meta name="keywords" content="<%=P("产品关键词")%>" />
<meta name="description" content="<%=P("产品描述")%>" />
<link href="<%=template%>css/style.css" rel="stylesheet" type="text/css">
<!--#include file="include/js.asp"-->
</head>
<body>
<!--#include file="top.asp"-->
<div id="mbody">
<div class="mbody-in">
<div class="breadcrumb"><a href=/><%=tw(10)%></a><span class="sep">/</span><%=P("产品标题")%></div>
<div class="sidebar">
<dl class="l_pro">
<dt><%=tw(8)%></dt>
<%=lmdh("产品分类")%>
</dl>
</div>

<!--内容详细{-->
    <div class="container">
        <h1 class="title"><%=P("产品标题")%></h1>
        <div class="byline">
            <span class="datetime"><%=tw(24)%>：<%=P("发布时间")%></span>
            <span class="views"><%=tw(23)%>：<%=P("点击次数")%></span>
        </div>
       
        
        <div align=center class="contentimg">
       <%=SHOWTPnr(1,2,5,60,500,400)%>
        </div>
   
        <div class="content">
          <%=P("产品内容")%>
        </div>

      
        <div class="prev-next">
            <div class="prev"><%=tw(18)%>：<%=P("上条信息")%></div>
            <div class="next"><%=tw(19)%>：<%=P("下条信息")%></div>
        </div>
      


</div>
</div>
</div>



</div>
</div>
<!--#include file="bottom.asp"-->

</body>
</html>
