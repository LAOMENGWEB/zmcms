
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=title%>--<%=tw(70)%>--<%=tw(98)%></title>
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

<dt>-<%=tw(71)%></dt>
<dd><a href="<%=sitepath%>member/memberinfo/?m=<%=request.QueryString("m")%>&language=<%=language%>"><%=tw(72)%></a></dd>
<dd><a href="<%=sitepath%>member/memberguest/?m=<%=request.QueryString("m")%>&language=<%=language%>"><%=tw(73)%></a></dd>
<dd><a href="<%=sitepath%>member/memberorder/?m=<%=request.QueryString("m")%>&language=<%=language%>"><%=tw(74)%></a></dd>
<dd><a href="<%=sitepath%>member/memberjob/?m=<%=request.QueryString("m")%>&language=<%=language%>"  <% if request.QueryString("m")=8 then response.write("class=""select""") end if%>><%=tw(75)%></a></dd>
<dd><a href="<%=sitepath%>inc/MemberLogin/?Action=Out&m=<%=request.QueryString("m")%>&language=<%=language%>"><%=tw(76)%></a></dd>





</dl>
<dl class="l_content">
<dt>-<%=tw(4)%>：<%=lxdz%></dt>
</dl>
<dl class="l_content">
<dt>-<%=tw(5)%>：<%=sj%></dt>
</dl>
<dl class="l_content">
<dt>-<%=tw(6)%>：<%=cz%></dt>
</dl>
<dl class="l_content">
<dt>-<%=tw(7)%>：<%=lxr%></dt>
</dl>

</div>

<!--右侧-->
<div class="fyRight">
<div class="title"> <span class="fl"><%=tw(70)%></span>
<div class="fr"><%=tw(10)%>：<%=tw(70)%> >> <%=tw(99)%></div>
</div>


<div class="mainRightMain">
<!--会员中心-应聘详细信息开始-->


<table border="0" cellpadding="0" cellspacing="0"  class="jobview">
  <tr>
    <td colspan="2" class="text"><%=tw(267)%></td>
  </tr>
  <tr>
    <td class="zkd"><%=tw(268)%></td>
    <td style="border-right:none"><div class=jlz><%=ypzw%></div></td>
  </tr>
  <tr>
    <td class="zkd"><%=tw(269)%></td>
    <td style="border-right:none"><div class=jlz><%=ypxm%></div></td>
  </tr>
  <tr>
    <td class="zkd"><%=tw(270)%></td>
    <td style="border-right:none"><div class=jlz><%=ypxb%></div></td>
  </tr>
  <tr>
    <td class="zkd"><%=tw(271)%></td>
    <td style="border-right:none"><div class=jlz><%=ypsg%>cm</div></td>
  </tr>
  <tr>
    <td class="zkd"><%=tw(272)%></td>
    <td style="border-right:none"><div class=jlz><%=yphj%></div></td>
  </tr>
  <tr>
    <td class="zkd"> <%=tw(273)%></td>
    <td style="border-right:none"><div class=jlz><%=yplxdh%></div></td>
  </tr>
  <tr>
    <td class="zkd"><%=tw(274)%></td>
    <td style="border-right:none"><div class=jlz><%=ypcsrq%></div></td>
  </tr>
  <tr>
    <td class="zkd"><%=tw(275)%></td>
    <td style="border-right:none"><div class=jlz><%=yphyzk%></div></td>
  </tr>
  <tr>
    <td class="zkd"><%=tw(276)%></td>
    <td style="border-right:none"><div class=jlz><%=yplxdz%></div></td>
  </tr>
  <tr>
    <td class="zkd"><%=tw(277)%></td>
    <td style="border-right:none"><div class=jlz><%=ypyzbm%></div></td>
  </tr>
  <tr>
    <td class="zkd"><%=tw(278)%></td>
    <td style="border-right:none"><div class=jlz><%=ypdzyj%></div></td>
  </tr>
  <%if yphfnr<>"" then %>
  <tr>
    <td class="zkd"><%=tw(279)%></td>
    <td style="border-right:none"><div class=jlz><%=yphfsj%> </div></td>
  </tr>
 
  <tr>
    <td class="zkd" style="border-bottom:none"><%=tw(280)%></td>
    <td style="border-bottom:none;border-right:none"><div class=jlz><%=HtmlStrReplace(yphfnr)%> </div></td>
  </tr>
  <%else%>
  <tr>
    <td class="zkd" style="border-bottom:none"><%=tw(280)%></td>
    <td style="border-bottom:none;border-right:none"><div class=jlz><%=tw(281)%></div></td>
  </tr>
  <%end if%>
</table>



<!--会员中心-应聘详细信息结束-->	
</div>
</div>

</div>
</div>




<!--#include file="../bottom.asp"-->
</body>
</html>
