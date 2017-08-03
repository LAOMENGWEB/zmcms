
<%=pfgg%>
<%=dlgg%>
<%=showzxkf%>
<%=share%>
<%=wzgb%>

<div class="top">
<div class="top1">
<div class="logo"><%=logoimg%></div>
<div class="zj"></div>
<div class="so">
<div class="language"><%=lan("多语言",1,5,"no","id asc")%></div>
<%=search(1,"搜 索")%> 
</div>
</div>
</div>
</div>
<div class="topmenu">
<%=lmdh("顶部导航标签")%>
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
