<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=tw(62)%>-<%=mc%>-<%=ym%></title>
<meta name="keywords" content="<%=gjc%>" />
<meta name="description" content="<%=ms%>" />
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

<dt><%=tw(63)%></dt>

<dd>
<a href="<%=sitepath%>guest/?m=9" title="<%=tw(64)%>" ><%=tw(64)%></a>
</dd>

<dd><a href="<%=sitepath%>showguest/?m=9" title="<%=tw(65)%>"  <% if request.QueryString("m")=9 then response.write("class=""select""") end if%>><%=tw(65)%></a></dd>





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
<div class="title"> <span class="fl"><%=tw(62)%></span>
<div class="fr"><%=tw(10)%>：<%=tw(66)%></div>
</div>


<div class="mainRightMain">


<%=infolist(7,1,"列表页留言","no","","","","",100,"yyyy/mm/dd","",100,"id desc")%>

</div>
</div>

</div>
</div>

<!--#include file="bottom.asp"-->



</body>
</html>