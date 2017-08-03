
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
<div class="fr"><%=tw(10)%>：<%=tw(70)%> >> <%=tw(101)%></div>
</div>


<div class="mainRightMain">

<!--会员中心-我的订单详情 代码开始-->
    
<input type="hidden" value="<%=isdate1%>" id="o_status" />
<div class="stepInfo">
<ul>
<li></li>
<li></li>
</ul>
<div class="stepIco stepIco1" id="create">1
<div class="stepText" id="createText"><%=tw(102)%></div>
</div>
<div class="stepIco stepIco2" id="check">2
<div class="stepText" id="checkText"><%=tw(103)%></div>
</div>
<div class="stepIco stepIco3" id="produce">3
<div class="stepText" id="produceText"><%=tw(104)%></div>
</div>
<div class="stepIco stepIco4" id="delivery">4
<div class="stepText" id="deliveryText"><%=tw(105)%></div>
</div>
<div class="stepIco stepIco5" id="received">5
<div class="stepText" id="receivedText"><%=tw(106)%></div>
</div>
</div>
    
    
    
    
    
    
    
    
    
    
<table  class="memberorderview" cellspacing="0" cellpadding="0">
<tr class="viewwz">
<td width="16%" ><%=tw(236)%></td>
<td width="26%" ><%=tw(237)%></td>
<td width="23%" ><%=tw(238)%></td>
<td width="10%" ><%=tw(239)%></td>
<td width="14%"><%=tw(240)%></td>
<td width="11%" style="border-right:none"><%=tw(241)%></td>
</tr>

<%=memberorderview("会员中心我的订单详细","no","yyyy-mm-dd","yyyy-mm-dd","id desc")%>
</table>


<table cellspacing="0" cellpadding="0" class="memberprderxx">
  <tr>
    <td colspan="2" class="text"><%=tw(252)%></td>
  </tr>
  <tr>
    <td class="zkd"><%=tw(253)%></td>
    <td style="border-right:none"><div class="jlz"><%=orderdgsj%></div></td>
  </tr>
    <tr>
    <td class="zkd"><%=tw(254)%></td>
    <td style="border-right:none"><div class="jlz"><%=orderzffs%></div></td>
  </tr>
    <tr>
    <td class="zkd"><%=tw(255)%></td>
    <td style="border-right:none"><div class="jlz"><%=orderpffs%> (<%=tw(256)%> <%=orderpsfy%> <%=tw(257)%>)</div></td>
  </tr>
      <tr>
    <td class="zkd"><%=tw(258)%></td>
    <td style="border-right:none;"><div class="jlz"><%=orderzfze%>  <%=tw(257)%>

<%if orderszzt=0 then%>
<a href="<%=orderzflj%>" onClick="return confirm('<%=tw(264)%> <%=orderzfze%> <%=tw(265)%>')"><font color="red"><%=tw(266)%></font></a>
<%end if%>

</div></td>
  </tr>
    <tr>
    <td class="zkd"><%=tw(259)%></td>
    <td style="border-right:none"><div class="jlz">  
	<%=orderddzt%></div></td>
  </tr>
    <tr>
    <td class="zkd" style="border-bottom:none"><%=tw(260)%></td>
    <td style="border-right:none;border-bottom:none">
	<div class="jlz">
	<%if orderhfnr<>"" then
	response.Write(""&tw(261)&""&orderclsj&""&tw(262)&"<br/>")
	response.Write(""&orderhfnr&"")
	else
	response.Write(""&tw(263)&"")
	end if
	%>
	
	</div>
	</td>
  </tr>
</TABLE>
  <table cellspacing="0" cellpadding="0" class="memberprderxx" style="margin-top:10px">
    <tr>
    <td colspan="2" class="text"><%=tw(242)%></td>
  </tr>
  <tr>
    <td class="zkd"><%=tw(243)%></td>
    <td style="border-right:none"><div class="jlz"><%=orderdgyh%></div></td>
  </tr>
  <tr>
    <td class="zkd"><%=tw(244)%></td>
    <td style="border-right:none"><div class="jlz"><%=orderyhxb%></div></td>
  </tr>
  <tr>
    <td class="zkd"><%=tw(245)%></td>
    <td style="border-right:none"><div class="jlz"><%=orderyhdw%></div></td>
  </tr>
   <tr>
    <td class="zkd"><%=tw(246)%></td>
    <td style="border-right:none"><div class="jlz"><%=orderyhdz%></div></td>
  </tr>
  <tr>
    <td class="zkd"><%=tw(247)%></td>
    <td style="border-right:none"><div class="jlz"><%=orderyhyb%></div></td>
  </tr>
  <tr>
    <td class="zkd"><%=tw(248)%></td>
    <td style="border-right:none"><div class="jlz"><%=orderdhhm%></div></td>
  </tr>
  <tr>
    <td class="zkd"><%=tw(249)%></td>
    <td style="border-right:none"><div class="jlz"><%=orderczhm%></div></td>
  </tr>
  <tr>
    <td class="zkd"><%=tw(250)%></td>
    <td style="border-right:none"><div class="jlz"><%=orderdzyx%></div></td>
  </tr>
  <tr>
    <td class="zkd" style="border-bottom:none"><%=tw(251)%></td>
    <td style="border-right:none;border-bottom:none"><div class="jlz"><%=orderddbz%></div></td>
  </tr> 
 
</table>

  








	<!--会员中心-我的订单详情 代码结束-->

</div>
</div>

</div>
</div>




<!--#include file="../bottom.asp"-->
</body>
</html>