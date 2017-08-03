<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=mc%>-<%=ym%></title>
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
	<dt><%=tw(48)%></dt>
<%=info(1,1,"单页分类","zmone="&zid&"",1,8,"","","",30,100,"yy","","id asc")%>
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
<div class="title"> <span class="fl"><%=tw(47)%></span>
<div class="fr"><%=tw(10)%>：<%=J("职位名称")%></div>
</div>
<div class="mainRightMain">



<div id="showjobsx">					
<ul class="paralist"> 
<li><span class="name"><%=tw(49)%>:</span><%=J("职位名称")%></li>              
<li><span class="name"><%=tw(50)%>:</span><%=J("需求人数")%></li>
<li><span class="name"><%=tw(51)%>:</span><%=J("用人单位")%></li>
<li><span class="name"><%=tw(52)%>:</span><%=J("工作地点")%></li>
<li><span class="name"><%=tw(53)%>:</span><%=J("工资")%></li>
<li><span class="name"><%=tw(54)%>:</span><%=J("单位地址")%></li>
<li><span class="name"><%=tw(55)%>:</span><%=J("联系电话")%></li>
<li><span class="name"><%=tw(56)%>:</span><%=J("投递邮箱")%></li>
<li><span class="name"><%=tw(57)%>:</span><%=J("点击次数")%></li>
<li><span class="name"><%=tw(58)%>:</span><%=J("招聘状态")%></li>  
<li><span class="name"><%=tw(59)%>:</span><a href="<%=J("应聘地址")%>"><%=tw(60)%></a></li>           
	   </ul>	</div>	




<div class="showjobsmbh">
<div class="showjobsm">
<h3 class="jobtitle"><span><%=tw(45)%></span></h3>
<%=J("详细说明")%>
</div>

<div class="jobnext">
<ul>
<li><%=tw(18)%>：<%=J("上条信息")%></li>
<li><%=tw(19)%>：<%=J("下条信息")%></li>
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
