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
<div class="fr">您当前的位置：营销网点查询</div>
</div>


<div class="mainRightMain">
<%=yxwlmap%>

<%=province("营销网络",5,"no","id desc")%>
</div>
</div>

</div>
</div>
<!--#include file="bottom.asp"-->


</body>
</html>