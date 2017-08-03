<!DOCTYPE HTML>
<html>
<head>
<meta content='text/html; charset=utf-8' http-equiv='Content-Type' />
<meta charset='utf-8' />
<title><%=T("视频标题")%>-<%=mc%>-<%=ym%></title>
<meta content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=0" name="viewport" />
<meta content="yes" name="apple-mobile-web-app-capable" />
<meta content="black" name="apple-mobile-web-app-status-bar-style" />
<meta content="telephone=no" name="format-detection" />
<link rel="stylesheet" type="text/css" href="<%=template%>wap/css/default.css" />
<link rel="stylesheet" href="<%=template%>wap/awesome/css/font-awesome.min.css">
<link rel="stylesheet" type="text/css" href="<%=template%>wap/slider/slider.css" />
<link rel="stylesheet" type="text/css" href="<%=template%>wap/slider/slider.default.css" /> 
<!--#include file="../../../wap/include/js.asp"-->
<!--播放器开始-->
<SCRIPT LANGUAGE=JavaScript>
<!--
function getLight(pars)
{	
	//alert("播放器模式参数值："+string+"");
	if(pars == "open")
	{
		close_light(this);
	}
	else
	{
		close_light(this);
	}
}


//兼容性
function thisMovie(movieName) {
    if (navigator.appName.indexOf("Microsoft") != -1) {
        return window[movieName]
    }
    else {
        return document[movieName]
    }
}

//-->
</SCRIPT> 
<!--播放器结束-->
</head>
<body>
<div class="zmcont">
<!--#include file="top.asp"-->

<header>
<div class="top mtop">
<div class="mreturn"><a href="javascript:history.go(-1);"></a></div>
<div class="mtitle"><%=tw(187)%></div>
<div class="rightnav"></div>
</div>
</header>


<div id="shownews">
<h1 class="title"><%=T("视频标题")%><p class="hits"><%=T("发布时间")%><%=tw(188)%>：<%=T("点击次数")%></p></h1>


<!--播放器开始-->
<div style="margin:10px; background:#000000">
<DIV class="close_light_bg" id="close_light_bg"></DIV>
<div class="video" id="CuPlayer">
<SCRIPT LANGUAGE=JavaScript>
<!--
var vID        = ""; 
var vWidth     = "100%";
var vHeight    = 200;
var vFile      = "CuSunV2set.xml";
var vPlayer    = "player.swf?v=2.5";
var vMp4url    = "../..<%=rsvideo("wapspurl")%>";
var vPic       = "../../<%=rsvideo("SmallImage")%>";
 
//-->
</SCRIPT> 
<script class="CuPlayerVideo" data-mce-role="CuPlayerVideo" type="text/javascript" src="Images/AsComV3.js"></script>
<SCRIPT language=javascript src="Images/action.js" type=text/javascript></SCRIPT>

</div>
</div>

<!--播放器结束-->
<div class="view"><%=tw(189)%></div>

<!--视频详情-->
<div class="editor" id="Gallery1">



<%=T("wap视频内容")%>

</div>
<div class="clear"></div>
</div>
<!--上下条-->
<div class="zm_page">
<ul class='preul'>
<li class='preinfo'>
<p class='pret'><%=tw(158)%></p>
<p class='pres'><%=T("WAP上条信息")%></p>
</li>
<li class='nextinfo'>
<p class='pret'><%=tw(159)%></p>
<p class='pres'><%=T("WAP下条信息")%></p>
</li>
</ul>
<div class="clear"></div>
</div>	

<!--#include file="bottom.asp"--> 
</div>


<script src="<%=sitepath%>wap/js/fun.inc.js" type="text/javascript"></script>
</body>
</html>