<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="admin.asp"-->
<!--#include file="conn.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>zmcms_管理页面</title>
<link href="css/css.css" rel="stylesheet" type="text/css" />
<script type="text/javascript" src="js/jquery-1.7.2.min.js"></script>
<script type="text/javascript">
$(document).ready(function(){
	
	$(".def-nav,.info-i").hover(function(){
		$(this).find(".pulldown-nav").addClass("hover");
		$(this).find(".pulldown").show();
	},function(){
		$(this).find(".pulldown").hide();
		$(this).find(".pulldown-nav").removeClass("hover");
	});
	
});
</script>
</head>
<body>
<div class="hd-main clearfix" id="header">
	
	<span class="logo"></span>
	
	<div class="navs">
		<a class="def-nav" href="set.asp" target="main">系统</a>
		<span class="separate"></span>
		<div class="def-nav  has-pulldown-special">
			<a class="pulldown-nav" href="danye.asp" target="main"><%=danye%></a>
		</div>
		
		<span class="separate"></span>
		
		<div class="def-nav  has-pulldown-special">
			<a class="pulldown-nav" href="xinwen.asp" target="main"><%=xinwen%></a>
			
		</div>
		
		<span class="separate"></span>
		<div class="def-nav  has-pulldown-special">
			<a class="pulldown-nav" href="chanpin.asp" target="main"><%=chanpin%></a>
			
		</div>
		
		<span class="separate"></span>
		<div class="def-nav  has-pulldown-special">
			<a class="pulldown-nav" href="anlie.asp" target="main"><%=tupian%></a>
			
		</div>
		
		<span class="separate"></span>
		<div class="def-nav  has-pulldown-special">
			<a class="pulldown-nav" href="shipin.asp" target="main"><%=shipin%></a>
			
		</div>
		
		<span class="separate"></span>
		<div class="def-nav  has-pulldown-special">
			<a class="pulldown-nav" href="xiazai.asp" target="main"><%=xiazai%></a>
			
		</div>
		
		<span class="separate"></span>
		<div class="def-nav  has-pulldown-special">
			<a class="pulldown-nav" href="zhaopin.asp" target="main"><%=zhaopin%></a>
			
		</div>
		
		<span class="separate"></span>
		<div class="def-nav  has-pulldown-special">
			<a class="pulldown-nav" href="huiyuan.asp" target="main">会员</a>
			
		</div>
		
		<span class="separate"></span>
		<div class="def-nav  has-pulldown-special">
			<a class="pulldown-nav" href="zhanghao.asp" target="main">账号</a>
			
		</div>
		
		<span class="separate"></span>
			<div class="def-nav  has-pulldown-special">
			<a class="pulldown-nav" href="anquan.asp" target="main">安全</a>
			
		</div>
		
		<span class="separate"></span>
		<a class="def-nav" href="chajian.asp" target="main">插件</a>
		<span class="separate"></span>
		<a class="def-nav" href="weixin.asp" target="main">微信</a>

	



		
	</div>
	
	<div class="info">
		<ul>
		
			<li class="info-i default-text">
				<a class="notice" href="../" target="_blank">首页预览</a>
		
			</li>
			<li class="info-i no-separate default-text has-pulldown">
				
				<span class="more" hideFocus="hideFocus"><a href="loginout.asp" target="_top">退出</a></span>
			
			</li>
		</ul>
	</div>
	
</div>

</body>
</html>
