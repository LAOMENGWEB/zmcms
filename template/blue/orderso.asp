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

<div class="ncbl">
<div class="d"><h2 class="title">关于我们</h2></div>
<div class="n"><%=info(1,1,"单页分类","zmone=1",1,8,"","","",30,100,"yy","","id asc")%></div>
<div class="f"></div>
<div class="d"><h2 class="title">联系我们</h2></div>
<div class="n"><div class="lx">

地址：<%=lxdz%><br />
电话：<%=sj%><br />
传真：<%=cz%><br />
联系人：<%=lxr%><br />

</div></div>
<div class="f"></div>

</div>
<div class="ncblr">
<div class="d"><h2 class="title">搜索结果</h2></div>
<div class="n"><div class="gsjj">

<%=showddso("订单结果显示","no","yy","yy","id asc")%>


	</div></div>
<div class="f"></div>
</div>
</div>
<!--#include file="bottom.asp"-->
<div class="clear"></div>
</div>


</body>
</html>


