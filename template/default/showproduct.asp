<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=P("产品标题")%><%=mc%>-<%=ym%></title>
<meta name="keywords" content="<%=P("产品关键词")%>" />
<meta name="description" content="<%=P("产品描述")%>" />
<link href="<%=template%>css/style.css" rel="stylesheet" type="text/css">
<!--#include file="include/js.asp"-->
<%If Trim(RsPro("P_SpecificationName")) <> "" And Not isnull(Trim(RsPro("P_SpecificationName"))) Then %>
<script>
//加载公共

$(document).ready(function(){ 
progg();

}); 

</script>
<%end if%>
<script>

//加载公共
$(document).ready(function(){ 
hygroup();

}); 




//会员组价格
function hygroup()
{
$.ajax({
	type:"get",
	url:"<%=template%>ajax/hygroup.asp",
	data:"proid=<%=proid%>&language=<%=language%>&r="+Math.round(Math.random()*10000),
	beforesend:function(){$("#hyjg").html("载入中...");},
	error:function(){$("#hyjg").html("获取信息失败");},
	success:function(data){$("#hyjg").html(data);}
});
}



//产品规格
function progg()
{
$.ajax({
	type:"get",
	url:"<%=template%>ajax/progg.asp",
	data:"proid=<%=proid%>&language=<%=language%>&r="+Math.round(Math.random()*10000),
	beforesend:function(){$("#cpgg").html("载入中...");},
	error:function(){$("#cpgg").html("获取信息失败");},
	success:function(data){$("#cpgg").html(data);}
});
}
</script>


</head>
<body>
<!--#include file="top.asp"-->
<div class="wrap">
<div class="main">

<div class="ppbh">
<div class="pptp"><%=SHOWTPnr(1,2,5,60,400,380)%></div>
<div class="ppsm">
<ul>
<li><%=tw(22)%>：<%=P("产品标题")%> </li>  

<li><%=tw(23)%>：<%=P("点击次数")%> </li>
<li><%=tw(24)%>：<%=P("发布时间")%></li>
<li><%=tw(25)%>：<s><%=hbfh%><%=P("市场价格")%></s></li>
<li><%=tw(26)%>：<%=P("产品库存")%>(<%=tw(27)%>：<%=P("产品销量")%>


<%=tw(26)%>：<%=P("剩余库存")%>
)</li>
<li style="border-bottom:none"><span id="result_0001_<%=ProId%>" style="font-weight:bold;font-size:14px;color:#F63"><%=word(623)%><%=hbfh%><b id="hyjg"></b></span></li>
</ul>
<!--自定义属性开始-->
<%=ProductInfoItemValue%>
<!--自定义属性结束-->

<!--商品规格开始-->
<div class="addtocartbox">
<div class="choosesp">
<span id="cpgg">
</span>
</div>  
</div> 
<!--商品规格结束-->

<!--提示信息带购物按开始按钮-->
<div id="myDiv" style=" background:#ff0000">

</div> 
<span id="carttip">
<img src="<%=template%>img/btn-cart.png" id="add_to_my_cart_img"  onClick="add_2_my_cart();"/>
</span>
<!--提示信息带购物按钮结束-->     


<input type="hidden" id="id" readonly value="<%=proid%>"/>
<input type="hidden" id="quantity" readonly value="1"/>




<!--购物车提示结束-->

</div>
</div>

<div class="showppsmbh">
<div class="showppsm">
<!--选型卡-->
<div class="fade">
<ul class="tabs">
<li><a href="#con1"><span><%=tw(28)%></span></a></li>

<%=content_title%>

</ul>
<div class="items">

<%=content2%>

</div>


</div>
<!--选型卡-->
</div>
<div class="pptags"><%=ptags("标签：",1,"/")%></div>	
<div class="ppnext">
<ul>
<li><%=tw(18)%>：<%=P("上条信息")%></li>
<li><%=tw(19)%>：<%=P("下条信息")%></li>
</ul>
</div>	  
</div>



</div>
</div>
<!--#include file="bottom.asp"-->

</body>
</html>
