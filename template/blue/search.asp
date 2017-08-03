<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=mc%>-<%=ym%></title>
<meta name="keywords" content="<%=gjza%>" />
<meta name="description" content="<%=msa%>" />
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
	<dt><%=tw(48)%></dt>
<%=info(1,1,"单页分类","zmone="&zid&"","",8,"","","",30,100,"yy","","id asc")%>
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
<div class="title"> <span class="fl"><%=tw(120)%></span>
<div class="fr"><%=tw(10)%>：<%=tw(297)%></div>
</div>
<div class="mainRightMain">

<%if tt=1 then%>
<%=showso(30,1,"搜索新闻页样式","no","","yyyy-mm-dd","",100,"<div style=""text-align:center;margin:10px 0 0"">","</div>","style=""color:red;font-weight:bold""","id desc")%>
<%end if%>
<%if tt=2 then%>
<%=showso(30,1,"搜索产品页样式","no","","yyyy-mm-dd","",100,"<div style=""text-align:center;margin:10px 0 0"">","</div>","style=""color:red;font-weight:bold""","id desc")%>
<%end if%>
<%if tt=3 then%>
<%=showso(30,1,"搜索图片页样式","no","","yyyy-mm-dd","",100,"<div style=""text-align:center;margin:10px 0 0"">","</div>","style=""color:red;font-weight:bold""","id desc")%>
<%end if%>
<%if tt=4 then%>
<%=showso(30,1,"搜索视频页样式","no","","yyyy-mm-dd","",100,"<div style=""text-align:center;margin:10px 0 0"">","</div>","style=""color:red;font-weight:bold""","id desc")%>
<%end if%>
<%if tt=5 then%>
<%=showso(30,1,"搜索下载页样式","no","","yyyy-mm-dd","",100,"<div style=""text-align:center;margin:10px 0 0"">","</div>","style=""color:red;font-weight:bold""","id desc")%>
<%end if%>

</div>

</div>
</div>
</div>


<div>
</div>
<!--#include file="bottom.asp"-->

</body>
</html>


