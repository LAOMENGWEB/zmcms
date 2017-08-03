<!DOCTYPE HTML>
<html>
<head>
<meta content='text/html; charset=utf-8' http-equiv='Content-Type' />
<meta charset='utf-8' />
<title><%=J("招聘职位")%>-<%=mc%>-<%=ym%></title>
<meta content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=0" name="viewport" />
<meta content="yes" name="apple-mobile-web-app-capable" />
<meta content="black" name="apple-mobile-web-app-status-bar-style" />
<meta content="telephone=no" name="format-detection" />
<link rel="stylesheet" type="text/css" href="<%=template%>wap/css/default.css" />
<link rel="stylesheet" href="<%=template%>wap/awesome/css/font-awesome.min.css">
<link rel="stylesheet" type="text/css" href="<%=template%>wap/slider/slider.css" />
<link rel="stylesheet" type="text/css" href="<%=template%>wap/slider/slider.default.css" /> 
<!--#include file="../../../wap/include/js.asp"-->
</head>
<body>
<div class="zmcont">
<!--#include file="top.asp"-->

<header>
<div class="top mtop">
<div class="mreturn"><a href="javascript:history.go(-1);"></a></div>
<div class="mtitle"><%=tw(162)%></div>
<div class="rightnav"></div>
</div>
</header>



<div id="showjob">
<div class="paralist">
<ul>
<li><%=tw(163)%>：<%=J("职位名称")%></li>              
<li><%=tw(164)%>：<%=J("需求人数")%></li>
<li><%=tw(165)%>：<%=J("用人单位")%></li>
<li><%=tw(166)%>：<%=J("工作地点")%></li>
<li><%=tw(167)%>：<%=J("工资")%></li>
<li><%=tw(168)%>：<%=J("单位地址")%></li>
<li><%=tw(169)%>：<%=J("联系电话")%></li>
<li><%=tw(170)%>：<%=J("投递邮箱")%></li>
<li><%=tw(171)%>：<%=J("点击次数")%></li>
<li><%=tw(172)%>：<%=J("招聘状态")%></li>  
</ul>
</div>
</div>

<div class="info_cv">
<div class="button">
<a href="<%=J("wap应聘地址")%>"><%=tw(173)%></a>
</div>
</div>

<div class="view"><%=tw(174)%></div>
<!--招聘详情-->
<div class="editor" id="Gallery1">


<%=J("wap详细说明")%>
</div>

<!--上下条-->
<div class="zm_page">
<ul class='preul'>
<li class='preinfo'>
<p class='pret'><%=tw(158)%></p>
<p class='pres'><%=j("WAP上条信息")%></p>
</li>
<li class='nextinfo'>
<p class='pret'><%=tw(159)%></p>
<p class='pres'><%=j("WAP下条信息")%></p>
</li>
</ul>
<div class="clear"></div>
</div>	

<!--#include file="bottom.asp"--> 
</div>

<script src="<%=sitepath%>wap/js/fun.inc.js" type="text/javascript"></script>
</body>
</html>