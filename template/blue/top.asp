<%=pfgg%>
<%=dlgg%>
<%=showzxkf%>
<%=share%>
<%=wzgb%>

<div id="header">
<div class="header-bg-line" id="header_bg_line"><div class="i"></div></div>
<div class="toper">
<div class="wrapper">
<div class="top-link">
<table><tr><td>
<div class="account-sign-status">
<span id="welcome"><%=tw(304)%><a href="<%=ym%>"><%=mc%></a>！</span>
</div>
</td><td><span class="sep">|</span></td><td>
<div class="top-nav">
<%=lan("首页顶部多语言",1,5,"no","id asc")%>
</div>
</td>
</tr></table>
</div>
</div>
</div>

<div class="header-line">
<div class="wrapper">
<div class="logo"><%=logoimg%></div>

<!--网站菜单开始-->
<div class="main-nav" id="main_nav">
<ul class="main">
<%=lmdh("顶部导航标签")%>      
</ul>
</div>
</div>
</div>   
</div>

<!--banner-->
<div class="topfoucs">
<div class="focusBox3">
<ul class="pic">
<%=diyfoucs("幻灯样式图片",3,"no","banner_order asc")%>
</ul>
<a class="prev" href="javascript:void(0)"></a>
<a class="next" href="javascript:void(0)"></a>
<ul class="hd">
<%=diyfoucs("幻灯样式数字",3,"no","banner_order asc")%>
</ul>
</div>
<script type="text/javascript">jQuery(".focusBox3").hover(function(){ jQuery(this).find(".prev,.next").stop(true,true).fadeTo("show",0.2) },function(){ jQuery(this).find(".prev,.next").fadeOut() });jQuery(".focusBox3").slide({ mainCell:".pic",effect:"fold", autoPlay:true, delayTime:600, trigger:"click"});</script></div>

</div>
