<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=mc%>-<%=ym%></title>
<meta name="keywords" content="<%=gjc%>" />
<meta name="description" content="<%=ms%>" />
<link href="<%=template%>css/style.css" rel="stylesheet" type="text/css">
<!--#include file="include/js.asp"-->
</head><div id="" class="">
	
</div>
<body >
<!--#include file="top.asp"-->
<div class="wrap"><div id="" class="">
	
</div>
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
   
<!--右侧-->
<div class="fyRight">
<div class="title"> <span class="fl"><%=tw(32)%></span>
<div class="fr"><%=tw(10)%>：<%=ddwz(4," >> ")%></div>
</div>


<div class="mainRightMain">
<% if request.QueryString("logdate")<>"" then%>
<%=guidanginfolist(4,1,0,"列表页视频",5,"no","","","",100,"yyyy-mm-dd","",100,"id desc")%>
<%else%>
<%=infolist(4,1,"列表页视频","no","","","","",100,"d","",100,"id desc")%>
<%end if%>
</div>
</div>

</div>
</div>

<!--#include file="bottom.asp"-->


</body>
</html>