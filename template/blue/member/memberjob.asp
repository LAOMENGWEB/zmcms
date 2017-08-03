
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
<div class="fr"><%=tw(10)%>：<%=tw(70)%> >> <%=tw(98)%></div>
</div>


<div class="mainRightMain">
	<!--会员中心应聘信息列表代码开始-->

<table  border="0" cellpadding="0" cellspacing="0" class="memberjob">
<tr class="jobwz">
<td width="24%"  ><%=word(205)%></td>
<td width="13%"  ><%=word(206)%></td>
<td width="20%" ><%=word(207)%></td>
<td width="23%"  ><%=word(208)%></td>
<td width="10%" ><%=word(209)%></td>
 <td width="10%" style="border-right:none"><%=word(210)%></td>
</tr>

<%=memberyp("会员中心应聘","id desc")%>
			
</table>







	<!--会员中心应聘信息列表代码结束-->
</div>
</div>

</div>
</div>




<!--#include file="../bottom.asp"-->


</body>
</html>
