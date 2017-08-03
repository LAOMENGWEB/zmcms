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
<div id="mbody">
<div class="mbody-in">
<div class="breadcrumb"><a href=/><%=tw(10)%></a><span class="sep">/</span><%=T("视频标题")%></div>
<div class="sidebar">
<dl class="l_pro">
<dt><%=tw(31)%></dt>
<%=lmdh("视频分类")%>
</dl>
</div>
<div class="container">
        <h1 class="title"><%=T("视频标题")%></h1>
        <div class="byline">

<div class="mksx"> <%=tw(14)%>：<%=T("视频作者")%>  　　<%=tw(15)%>：<%=T("发布时间")%>  　　<%=tw(16)%>：<%=T("视频来源")%> 　　<%=tw(17)%>：<%=T("点击次数")%></div>
        </div>
       
        
        <div align=center class="contentimg">
      <%=SHOWTV%>
        </div>
   
        <div class="content">
        <%=T("视频内容")%>
        </div>

      
        <div class="prev-next">
            <div class="prev"><%=tw(18)%>：<%=T("上条信息")%></div>
            <div class="next"><%=tw(19)%>：<%=T("下条信息")%></div>
        </div>
       
    </div>

</div>
</div>
</div>





<!--#include file="bottom.asp"-->
</body>
</html>