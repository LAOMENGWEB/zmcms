<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=title%>-<%=tw(100)%></title>
<meta name="keywords" content="<%=gjc%>" />
<meta name="description" content="<%=ms%>" />
<link href="<%=template%>css/style.css" rel="stylesheet" type="text/css">
<!--#include file="../include/js.asp"-->
</head>
<body >
<!--#include file="../top.asp"-->
<div class="wrap">
<div class="main">
<!--左侧-->
<div class="fyLeft">
<dl class="l_pro">
<dt><%=tw(71)%></dt>

<dd><a href="<%=sitepath%>member/memberinfo/?m=<%=request.QueryString("m")%>&language=<%=language%>"><%=tw(72)%></a></dd>
<dd><a href="<%=sitepath%>member/memberguest/?m=<%=request.QueryString("m")%>&language=<%=language%>"><%=tw(73)%></a></dd>
<dd><a href="<%=sitepath%>member/memberorder/?m=<%=request.QueryString("m")%>&language=<%=language%>" <% if request.QueryString("m")=8 then response.write("class=""select""") end if%>><%=tw(74)%></a></dd>
<dd><a href="<%=sitepath%>member/memberjob/?m=<%=request.QueryString("m")%>&language=<%=language%>"><%=tw(75)%></a></dd>
<dd><a href="<%=sitepath%>inc/MemberLogin/?Action=Out&m=<%=request.QueryString("m")%>&language=<%=language%>"><%=tw(76)%></a></dd>
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
<div class="title"> <span class="fl"><%=tw(70)%></span>
<div class="fr"><%=tw(10)%>：<%=tw(70)%> >> <%=tw(100)%></div>
</div>


<div class="mainRightMain">
	<!--会员中心-我的订单 列表开始-->

<table  border="0" cellpadding="0" cellspacing="0" class="memberorder">
<tr class="orderwz"n>
<td width="17%" ><%=tw(230)%></td>
<td width="10%"  ><%=tw(231)%></td>
<td width="12%" ><%=tw(232)%></td>
<td width="18%"  ><%=tw(233)%></td>
<td width="9%" ><%=tw(234)%></td>
<td width="26%" style="border-right:none"><%=tw(235)%></td>
 </tr>
     
   
	<%=membershowdd("会员中心我的订单","no",1,"yyyy-mm-dd","yyyy-mm-dd","id desc")%>		

</table>

	<!--会员中心-我的订单 列表结束-->
</div>
</div>

</div>
</div>




<!--#include file="../bottom.asp"-->


</body>
</html>