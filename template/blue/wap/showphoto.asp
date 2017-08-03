<!DOCTYPE HTML>
<html>
<head>
<meta content='text/html; charset=utf-8' http-equiv='Content-Type' />
<meta charset='utf-8' />
<title><%=C("图片标题")%>-<%=mc%>-<%=ym%></title>
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
<div class="mtitle"><%=tw(175)%></div>
<div class="rightnav"></div>
</div>
</header>


<div id="shownews">
<h1 class="title"><%=C("图片标题")%><p class="hits"><%=C("发布时间")%><%=tw(176)%>：<%=C("点击次数")%></p></h1>


<!--多图展示开始-->
<div class="pshow">
<div id="focus1" class="focus1">
<div class="hd">
<ul></ul>
</div>
<div class="bd">
<ul id="Gallery">
<%
for i=0 to ubound(C("图片组"))
%>
<a href="../..<%=C("图片组")(i)%>"><li><img _src="../..<%=C("图片组")(i)%>"  src="../images/blank.png" /></li></a>
<%
next
%>
</ul>
</div>
</div>
</div>
<script type="text/javascript">
TouchSlide({ 
slideCell:"#focus1",
titCell:".hd ul", //开启自动分页 autoPage:true ，此时设置 titCell 为导航元素包裹层
mainCell:".bd ul", 
effect:"left", 
autoPlay:true,//自动播放
autoPage:true, //自动分页
switchLoad:"_src" //切换加载，真实图片路径为"_src" 
});
</script>
<!--多图展示结束-->
<div style="padding-top:20px">
<div class="view"><%=tw(177)%></div>
</div>
<!--图片详情-->
<div class="editor" id="Gallery1">



<%=C("wap图片内容")%>

</div>
<div class="clear"></div>
</div>
<!--上下条-->
<div class="zm_page">
<ul class='preul'>
<li class='preinfo'>
<p class='pret'><%=tw(158)%></p>
<p class='pres'><%=C("WAP上条信息")%></p>
</li>
<li class='nextinfo'>
<p class='pret'><%=tw(159)%></p>
<p class='pres'><%=C("WAP下条信息")%></p>
</li>
</ul>
<div class="clear"></div>
</div>	

<!--#include file="bottom.asp"--> 
</div>
<script src="<%=sitepath%>wap/js/fun.inc.js" type="text/javascript"></script>
</body>
</html>