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
<%=info(1,1,"单页分类","zmone=1",1,8,"","","",30,100,"yy","","id asc")%>
</dl>
<dl class="l_content">
<dt>地址：<%=lxdz%></dt>
</dl>
<dl class="l_content">
<dt>电话：<%=sj%></dt>
</dl>
<dl class="l_content">
<dt>传真：<%=cz%></dt>
</dl>
<dl class="l_content">
<dt>负责：<%=lxr%></dt>
</dl>

</div>
   
<!--右侧-->
<div class="fyRight">
<div class="title"> <span class="fl">营销网络</span>
<div class="fr">您当前的位置：营销网点详细信息</div>
</div>


<div class="mainRightMain">

<div id="showjobsx">					
<ul class="paralist"> 
<li><span class="name">销售网点：</span><%=Y("所属省份")%></li>              
<li><span class="name">公司名称：</span><%=Y("公司名称")%></li>
<li><span class="name">公司法人：</span><%=Y("联系人")%></li>
<li><span class="name">联系地址：</span><%=Y("联系地址")%></li>
<li><span class="name">联系电话：</span><%=Y("联系电话")%></li>
<li><span class="name">传真号码：</span><%=Y("传真号码")%></li>
<li><span class="name">公司网址：</span><%=Y("公司网址")%></li>
<li><span class="name">邮政编码：</span><%=Y("邮政编码")%></li>
        
	   </ul>	</div>	

<div class="showjobsmbh">
<div class="showjobsm">
<h3 class="jobtitle"><span>详细描述</span></h3>
<%=Y("详细说明")%>
</div>

<div class="jobnext">
<ul>
<li>上条信息：<%=Y("上条信息")%></li>
<li>下条信息：<%=Y("下条信息")%></li>
</ul>
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
