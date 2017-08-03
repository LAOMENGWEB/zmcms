<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../../../Inc/sub.asp" -->  
<%
'初始化加密类
call initCodecs 

orderddh=base64Decode(request.QueryString("orderddh"))
action=request.QueryString("action")
zffs=request.QueryString("zffs")
if action="gg" then
if zffs=1 then
sql4="update orderdd set zffs=2 where orderddh='"&orderddh&"' and Lang='"&Language&"'"
conn.Execute ( sql4 )
end if
if zffs=2 then
sql4="update orderdd set zffs=1 where orderddh='"&orderddh&"' and Lang='"&Language&"'"
conn.Execute ( sql4 )
end if
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=title%>-<%=tw(67)%></title>
<meta name="keywords" content="<%=gjc%>" />
<meta name="description" content="<%=ms%>" />
<!--#include file="../include/js.asp"-->
<link href="<%=template%>css/style.css" rel="stylesheet" type="text/css">
</head>
<body >
<%
IF ykgw=2 THEN
if session("MemName")="" or session("GroupID")="" or session("GroupID")=999177 or session("MemLogin")<>"Succeed" then
response.redirect ""&sitepath&"member/memberuserlogin/?m="&m2&"&Language="&Language&""
end if
end if
%>
<!--#include file="../top.asp"-->
<div class="wrap">
<div class="main">
<div id="cart_div">
<div class="cart_quick">
<div class="cart_ico"><%=word(532)%></div>
<div class="cart_step">
<ul>
<li class="step1"><span class="active"><span><span><%=word(533)%></span></span></span></li>
<li class="step2"><span class="active"><span><span><%=word(534)%></span></span></span></li>
<li class="step3"><span class="active"><span><span><%=word(535)%></span></span></span></li>
</ul>
</div>
</div>

<%
set rs=server.createobject("adodb.recordset")
sql="select * from orderdd where orderddh='"&orderddh&"' order by ID desc"
rs.open sql,conn,1,1
%>

<div id="cart_order">
<div id="cartshow">
<%if rs("zffs")=1 then%>
<%if alipayapi=1 then%>
<form name=alipayment action=<%=sitepath%>pay/jsdz/alipayapi.asp method=post target="_blank">
<%end if%>
<%if alipayapi=2 then%>
<form name=alipayment action=<%=sitepath%>pay/dbjy/alipayapi.asp method=post target="_blank">
<%end if%>
<%if alipayapi=3 then%>
<form name=alipayment action=<%=sitepath%>pay/sjk/alipayapi.asp method=post target="_blank">
<%end if%>
<%end if%>
<%if rs("zffs")=2 then%>
<form name=alipayment action=<%=sitepath%>pay/blank/?m=<%=m3%>&Language=<%=language%> method=post target="_blank">
<%end if%>
<table>
<tr>
<th colspan="2"><%=word(601)%></th>
</tr>
<tr>
<td class="info"><%=word(602)%></td>
<td><%=rs("orderddh")%><input name="WIDout_trade_no" value="<%=rs("orderddh")%>" size="30" type="hidden" /><a href="<%=sitepath%>member/memberorderview/?orderid=<%=rs("id")%>&m=<%=m2%>&language=<%=language%>&orderdd=<%=orderddh%>" target="_blank">  <%=word(603)%></a>   
</td>
</tr>
<tr>
<td class="info"><%=word(604)%></td>
<td><%=rs("shipmentName")%></td>
</tr>
<tr>
<td class="info"><%=word(605)%></td>
<td><span><%=formatNumber2(rs("aumont"))%> <%=hbwz%><input size="30" name="WIDsubject" value="<%=rs("ordername")%>" type="hidden" /></span></td>
</tr>
<tr>
<td class="info"><%=word(606)%></td>
<td><span><%=formatNumber2(rs("shipmentMoney"))%> <%=hbwz%></span> <input size="30" name="WIDbody" value="<%=rs("Remark")%>" type="hidden"/><input size="30" name="WIDshow_url" value="<%=ym%>/member/memberorderview/?id2=<%=rs("id")%>" type="hidden"/></td>
 </tr>
  <tr>
<td class="info"><%=word(607)%></td>
 <td><span><%=formatNumber2(rs("aumont")+rs("shipmentMoney"))%> <%=hbwz%></span><input size="30" name="WIDtotal_fee" value="<%=formatNumber2(rs("aumont")+rs("shipmentMoney"))%>" type="hidden"/><input size="30" name="WIDprice" value="<%=formatNumber2(rs("aumont")+rs("shipmentMoney"))%>" type="hidden"/><input size="30" name="WIDreceive_name" type="hidden" vaule="<%=rs("linkman")%>"/><input size="30" name="WIDreceive_address" type="hidden" vaule="<%=rs("address")%>"/><input size="30" name="WIDreceive_zip" type="hidden" vaule="<%=rs("zipcode")%>"/><input size="30" name="WIDreceive_mobile" vaule="<%=rs("mobile")%>" type="hidden"/></td>
  </tr>
  </table>
  <table>
<tr>
  <th colspan="2"><%=word(608)%></th>
</tr>
<tr>
<td class="info"><%=word(609)%></td>
 <td><%if rs("zffs")=1 then response.write(""&word(610)&"") end if%><%if rs("zffs")=2 then response.write(""&word(611)&"") end if%> <a href="?orderddh=<%=base64Encode(orderddh)%>&zffs=<%=rs("zffs")%>&action=gg&m=<%=request.QueryString("m")%>&Language=<%=language%>"> <%=word(612)%></a></td>
</tr>
<tr>
  <td><%=word(613)%></td>
  <td> <%
			  if rs("isdate")=0 then
			  response.write ""&word(614)&"" 
			  end if
			  if rs("isdate")=1 then  
			  response.write ""&word(615)&"" 
			  end if
			   if rs("isdate")=2 then  
			  response.write ""&word(616)&"" 
			  end if
			   if rs("isdate")=3 then  
			  response.write ""&word(617)&"" 
			  end if
			   if rs("isdate")=5 then  
			  response.write ""&word(618)&"" 
			  end if
			  %>
			  </td>
</tr>
<%if rs("isdate")=0 then%>
<tr>
<td colspan="2"><input type="submit" name="send" value="<%=word(619)%>" class="zfan" /><%if session("MemName")="" or session("GroupID")="" or session("GroupID")=999177 or session("MemLogin")<>"Succeed" then%>
 <%=word(620)%>
<%ELSE%>
<input type="button" value="<%=word(621)%>"  onClick="location.href='<%=sitepath%>member/memberorder/?m=<%=m2%>&Language=<%=language%>'" class="yhan"/><%end if%></td>
</tr>
<%end if%>
</table>
</form>
</div>          
</div>
</div>

<%
rs.close
set rs=nothing
%>

</div>
</div>

<!--#include file="../bottom.asp"-->
</body>
</html>