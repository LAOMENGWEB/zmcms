<!-- #include file="../../../Inc/sub.asp" --> 
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=mc%>-在线下单</title>
<meta name="keywords" content="<%=gjc%>" />
<meta name="description" content="<%=ms%>" />
<link href="<%=template%>css/style.css" rel="stylesheet" type="text/css">
<!--#include file="../include/js.asp"-->

</head>

<body>
<!--#include file="../top.asp"-->
<div class="wrap">
<div class="main" style="margin-bottom:40px">

<!--加入购物车开始-->
<%
step=request.QueryString("step")
sql="select * from product where ID="&Trim(Request("ID"))&" and Lang='"&Language&"'"
set rs=server.CreateObject("adodb.recordset")
rs.open sql,conn,1,1
if not rs.eof Then
myAction = Trim(Request("myAction"))
If Trim(Request("myAction")) = "" Then myAction = "addProduct"
Select Case myAction
Case "addProduct"
quantity = Trim(Request("quantity"))
name = Trim(rs("title"))
img=rs("SmallImage")
'判断是否有商品规格的产品价格S
If Trim(rs("P_SpecificationName")) <> "" And Not isnull(Trim(rs("P_SpecificationName"))) Then 
retailprice = Trim(rs("scprice")+rs("p_price"))
if session("MemName")="" or session("GroupID")="" or session("GroupID")=999177 or session("MemLogin")<>"Succeed"   then
price = Trim(rs("hyprice")+rs("p_price"))
else

set rsjg=server.CreateObject("adodb.recordset")
sqljg="select * from hygroup where groupid='"&session("GroupID")&"' and Lang='"&Language&"'"
rsjg.open sqljg,conn,1,1
set rsjg1=server.CreateObject("adodb.recordset")
sqljg1="select * from product_price where grade_id="&rsjg("GroupLevel")&" and product_id="&request.QueryString("ID")&" and Lang='"&Language&"'"
rsjg1.open sqljg1,conn,1,1
price = Trim(rsjg1("price")+rs("p_price"))

end if
else
retailprice = Trim(rs("scprice"))
if session("MemName")="" or session("GroupID")="" or session("GroupID")=999177 or session("MemLogin")<>"Succeed"   then
price = Trim(rs("hyprice"))
else
set rsjg=server.CreateObject("adodb.recordset")
sqljg="select * from hygroup where groupid='"&session("GroupID")&"' and Lang='"&Language&"'"
rsjg.open sqljg,conn,1,1
set rsjg1=server.CreateObject("adodb.recordset")
sqljg1="select * from product_price where grade_id="&rsjg("GroupLevel")&" and product_id="&request.QueryString("ID")&" and Lang='"&Language&"'"
rsjg1.open sqljg1,conn,1,1
price = Trim(rsjg1("price"))
end if
end if
'判断是否有商品规格的产品价格E
If quantity = "" Then
quantity = 1
End If
zhid=clng(rs("p_sjs"))
'判断是否有商品规格 有的话就添加 规格的ID值
If Trim(rs("P_SpecificationName")) <> "" And Not isnull(Trim(rs("P_SpecificationName"))) Then 
Call AddItem(clng(rs("p_id")),quantity,name,retailprice,price,img,""&rs("p_id")&"")
else
Call AddItem(clng(rs("p_sjs")),quantity,name,retailprice,price,img,0)
end if

'修改购物车
Case "modifyNumber"
sql2="select * from product_stock where ID="&Trim(Request("sid"))&" and Lang='"&Language&"'"
set rs2=server.CreateObject("adodb.recordset")
rs2.open sql2,conn,1,1
if not rs2.eof then
sql3="select * from product where ID="&rs2("p_id")&" and Lang='"&Language&"'"
set rs3=server.CreateObject("adodb.recordset")
rs3.open sql3,conn,1,1

sid = Trim(Request("sid"))
myNumber = Trim(Request("quantity"))
name = Trim(rs3("title"))
img=rs3("SmallImage")
retailprice = Trim(rs3("scprice")+rs2("p_price"))
if session("MemName")="" or session("GroupID")="" or session("GroupID")=999177 or session("MemLogin")<>"Succeed"   then
price = Trim(rs3("hyprice")+rs2("p_price"))
else
set rsjg=server.CreateObject("adodb.recordset")
sqljg="select * from hygroup where groupid='"&session("GroupID")&"' and Lang='"&Language&"'"
rsjg.open sqljg,conn,1,1
set rsjg1=server.CreateObject("adodb.recordset")
sqljg1="select * from product_price where grade_id="&rsjg("GroupLevel")&" and product_id="&rs2("p_id")&" and Lang='"&Language&"'"
rsjg1.open sqljg1,conn,1,1
price = Trim(rsjg1("price")+rs2("p_price"))

end if
Call UpdateItem(rs2("id"),CLng(myNumber),name,retailprice,price,img,""&Trim(Request("sid"))&"")
else
sql5="select * from product where p_sjs="&Trim(Request("sid"))&" and Lang='"&Language&"'"
set rs5=server.CreateObject("adodb.recordset")
rs5.open sql5,conn,1,1
myNumber = Trim(Request("quantity"))
name = Trim(rs5("title"))
img=rs5("SmallImage")
retailprice = Trim(rs5("scprice"))
if session("MemName")="" or session("GroupID")="" or session("GroupID")=999177 or session("MemLogin")<>"Succeed"   then
price = Trim(rs5("hyprice"))
else
set rsjg=server.CreateObject("adodb.recordset")
sqljg="select * from hygroup where groupid='"&session("GroupID")&"' and Lang='"&Language&"'"
rsjg.open sqljg,conn,1,1
set rsjg1=server.CreateObject("adodb.recordset")
sqljg1="select * from product_price where grade_id="&rsjg("GroupLevel")&" and product_id="&rs5("id")&" and Lang='"&Language&"'"
rsjg1.open sqljg1,conn,1,1
price = Trim(rsjg1("price"))
end if
Call UpdateItem(rs5("p_sjs"),CLng(myNumber),name,retailprice,price,img,0)

end if
'删除购物车中某1个产品
Case "del"
mmid = Trim(Request("mid")) 
Call RemoveItem(mmid)
%>

<%
'清空购物车
Case "clearCart"
Session("Cart") = Null
Session("Cart") = ""
%>

<%
End Select
%>
<%if step=1 then%>
<div id="cart_div">
<div class="cart_quick">
<div class="cart_ico"><%=word(532)%></div>
<div class="cart_step">
<ul>
<li class="step1"><span class="active"><span><span><%=word(533)%></span></span></span></li>
<li class="step2"><span><span><span><%=word(534)%></span></span></span></li>
<li class="step3"><span><span><span><%=word(535)%></span></span></span></li>
</ul>
</div>
</div>
<%
CheckCart = Request("CheckCart++")
If CheckCart = true Then
%>
<table  border="0" cellpadding="0" cellspacing="0" class="productgwc">
<form action=""  method=POST name=CartFoem>
<tr class="gwcwz">
<td width="9%" ><%=word(536)%></td>
<td width="30%"  ><%=word(537)%></td>
<td width="11%"  ><%=word(538)%></td>
<td width="9%" ><%=word(539)%></td>
<td width="15%" ><%=word(540)%></td>
<td width="17%" ><%=word(541)%></td>
<td width="9%" style="border-right:none"><%=word(542)%></td>
</tr>
<%
aryCart = Session("Cart")
i = 0
%>
<%
strTotalretailprice = 0
strTotalprice = 0
strcount = 0
For m = 0 to Ubound(aryCart,2)
%>
<tr class="gwcsj">
<td >
<%
if aryCart(i+6,m)=0 then
sql1="select * from product where p_sjs="&aryCart(i,m)&" and Lang='"&Language&"'"
set rs1=server.CreateObject("adodb.recordset")
rs1.open sql1,conn,1,1%>
<%if rs1("zmone")<>0 and rs1("zmtwo")=0 and rs1("zmthree")=0 then%>
<a href=<%=sitepath%>showproduct/?pone=<%=rs1("zmone")%>&proID=<%=rs1("id")%>>
<%end if%>
<%if rs1("zmone")<>0 and rs1("zmtwo")<>0 and rs1("zmthree")=0 then%>
<a href=<%=sitepath%>showproduct/?pone=<%=rs1("zmone")%>&ptwo=<%=rs1("zmtwo")%>&proID=<%=rs1("id")%>>
<%end if%>
<%if rs1("zmone")<>0 and rs1("zmtwo")<>0 and rs1("zmthree")<>0 then%>
<a href=<%=sitepath%>showproduct/?pone=<%=rs1("zmone")%>&ptwo=<%=rs1("zmtwo")%>&pthree=<%=rs1("zmthree")%>&proID=<%=rs1("id")%>>
<%end if%>
<%if aryCart(i+5,m)<>"" then%>
<img src="<%=sitepath%><%=aryCart(i+5,m)%>" width=80px height=40px/>
<%else%>
<img src="<%=sitepath%>images/demo.jpg" width=80px height=40px/>
<%end if%>
</a>

<%
set rs1=nothing
rs1.close

else


sql1="select * from product_stock where ID="&aryCart(i,m)&" and Lang='"&Language&"'"
set rs1=server.CreateObject("adodb.recordset")
rs1.open sql1,conn,1,1

sql2="select * from product where ID="&rs1("p_id")&" and Lang='"&Language&"'"
set rs2=server.CreateObject("adodb.recordset")
rs2.open sql2,conn,1,1

%>
<%if rs2("zmone")<>0 and rs2("zmtwo")=0 and rs2("zmthree")=0 then%>
<a href=<%=sitepath%>showproduct/?pone=<%=rs2("zmone")%>&proID=<%=rs2("id") %>>
<%end if%>
<%if rs2("zmone")<>0 and rs2("zmtwo")<>0 and rs2("zmthree")=0 then%>
<a href=<%=sitepath%>showproduct/?pone=<%=rs1("zmone")%>&ptwo=<%=rs1("zmtwo")%>&proID=<%=rs2("id") %>>
<%end if%>
<%if rs2("zmone")<>0 and rs2("zmtwo")<>0 and rs2("zmthree")<>0 then%>
<a href=<%=sitepath%>showproduct/?pone=<%=rs2("zmone")%>&ptwo=<%=rs2("zmtwo")%>&pthree=<%=rs2("zmthree")%>&proID=<%=rs2("id") %>>
<%end if%>
<%if aryCart(i+5,m)<>"" then%>
<img src="<%=sitepath%><%=aryCart(i+5,m)%>" width=80px height=40px/>
<%else%>
<img src="<%=sitepath%>images/demo.jpg" width=80px height=40px/>
<%end if%>
</a>

<%
set rs2=nothing
rs2.close

set rs1=nothing
rs1.close
%>
<%end if%>
<input name="chkpay" type="checkbox" style="display:none" value="<%=aryCart(i,m) %>" checked="checked" />	</td> 
<td>
<input type=hidden value=<%=aryCart(i,m) %> name=ID />
<%

if aryCart(i+6,m)=0 then
sql1="select * from product where p_sjs="&aryCart(i,m)&" and Lang='"&Language&"'"
set rs1=server.CreateObject("adodb.recordset")
rs1.open sql1,conn,1,1%>
<%if rs1("zmone")<>0 and rs1("zmtwo")=0 and rs1("zmthree")=0 then%>
<a href=<%=sitepath%>showproduct/?pone=<%=rs1("zmone")%>&proID=<%=rs1("id")%>>
<%end if%>
<%if rs1("zmone")<>0 and rs1("zmtwo")<>0 and rs1("zmthree")=0 then%>
<a href=<%=sitepath%>showproduct/?pone=<%=rs1("zmone")%>&ptwo=<%=rs1("zmtwo")%>&proID=<%=rs1("id")%>>
<%end if%>
<%if rs1("zmone")<>0 and rs1("zmtwo")<>0 and rs1("zmthree")<>0 then%>
<a href=<%=sitepath%>showproduct/?pone=<%=rs1("zmone")%>&ptwo=<%=rs1("zmtwo")%>&pthree=<%=rs1("zmthree")%>&proID=<%=rs1("id")%>>
<%
end if
set rs1=nothing
rs1.close

%>
<%=aryCart(i+2,m) %></a>
<%else%>

<%




sql1="select * from product_stock where ID="&aryCart(i,m)&" and Lang='"&Language&"'"
set rs1=server.CreateObject("adodb.recordset")
rs1.open sql1,conn,1,1

sql2="select * from product where ID="&rs1("p_id")&" and Lang='"&Language&"'"
set rs2=server.CreateObject("adodb.recordset")
rs2.open sql2,conn,1,1

%>
<%if rs2("zmone")<>0 and rs2("zmtwo")=0 and rs2("zmthree")=0 then%>
<a href=<%=sitepath%>showproduct/?pone=<%=rs2("zmone")%>&proID=<%=rs2("id") %>>
<%end if%>
<%if rs2("zmone")<>0 and rs2("zmtwo")<>0 and rs2("zmthree")=0 then%>
<a href=<%=sitepath%>showproduct/?pone=<%=rs1("zmone")%>&ptwo=<%=rs1("zmtwo")%>&proID=<%=rs2("id") %>>
<%end if%>
<%if rs2("zmone")<>0 and rs2("zmtwo")<>0 and rs2("zmthree")<>0 then%>
<a href=<%=sitepath%>showproduct/?pone=<%=rs2("zmone")%>&ptwo=<%=rs2("zmtwo")%>&pthree=<%=rs2("zmthree")%>&proID=<%=rs2("id") %>>
<%end if%>


<%=aryCart(i+2,m) %></a>

<%
' 购物车显示规格
sql3="select * from product_stock where ID="&aryCart(i,m)&" and Lang='"&Language&"' "
set rs3=server.CreateObject("adodb.recordset")
rs3.open sql3,conn,1,1

response.write("<br />"&word(543)&""&rs3("s_value")&" <br />"&word(544)&""&rs3("p_stock")&" ")

set rs3=nothing
rs3.close
%>




<%
end if
set rs2=nothing
rs2.close

set rs1=nothing
rs1.close
%>



<input type=hidden value=<%=aryCart(i+2,m) %> name=name<%=m+1 %>></td> 	
<td ><%=formatnumber2(aryCart(i+3,m) )%> <%=hbwz%><input type=hidden value=<%=formatnumber2(aryCart(i+3,m) )%> name=retailprice<%=m+1 %>></td> 
<td ><%=formatnumber2(aryCart(i+4,m)) %> <%=hbwz%>
<input type=hidden value=<%=formatnumber2(aryCart(i+4,m)) %> name=price<%=m+1 %> id=price<%=aryCart(i,m)%> /></td> 
<td >
<div  class="amount-box">
<span class="handler">
<a class="down" href="#minus" onClick="return product_minus(<%=aryCart(i,m)%>)" title="减少">减少</a></span>
<span class="enter">
<input type=text value=<%=aryCart(i+1,m) %> name="quantity"  onKeyUp="checkout(<%=aryCart(i,m)%>)" id="quantity<%=aryCart(i,m) %>" >
</span>
<span class="handler">

<%if aryCart(i+6,m)<>0 then
sql1="select * from product_stock where ID="&aryCart(i,m)&" and Lang='"&Language&"'"
set rs1=server.CreateObject("adodb.recordset")
rs1.open sql1,conn,1,1

kc=rs1("p_stock")

%>
<a class="up" href="#plus" onClick="return product_plus(<%=aryCart(i,m) %>,<%=kc%>) " title="增加"></a></span>

<%end if
set rs1=nothing
rs1.close

%>
<%if aryCart(i+6,m)=0 then
sql1="select * from product where p_sjs="&aryCart(i,m)&" and Lang='"&Language&"'"
set rs1=server.CreateObject("adodb.recordset")
rs1.open sql1,conn,1,1

kc=rs1("kucun1")

%>
<a class="up" href="#plus" onClick="return product_plus(<%=aryCart(i,m) %>,<%=kc%>) " title="增加"></a></span>

<%end if
set rs1=nothing
rs1.close

%>





</div>		
</td> 
<td ><%=hbfh%> <span id="money<%=aryCart(i,m) %>"><%=formatNumber2(aryCart(i+1,m)*aryCart(i+4,m)) %></span></td>

<td style="border-right:none"><a href=?step=1&myAction=del&mID=<%=aryCart(i,m)%>&m=<%=request.QueryString("m")%>&Language=<%=language%>><%=word(545)%></a></td></tr>
<%
strTotalretailprice=strTotalretailprice+aryCart(i+1,m)*aryCart(i+3,m)
strTotalprice=strTotalprice+aryCart(i+1,m)*aryCart(i+4,m)
strcount=strcount+aryCart(i+1,m)
Next
%>
</tr>
<tr class="gwcan">
<td colspan="8" style="border-bottom:none;border-right:none;" >
<span><%=word(546)%><b><label id="allpay"><%=hbfh%><%=formatNumber2(strTotalprice)%></label></b> <%=hbwz%></span> 
<a href="<%=sitepath%>product/?m=<%=request.QueryString("m")%>&Language=<%=language%>" class="go"><%=word(547)%></a>  <a href="?step=1&myAction=clearCart&m=<%=request.QueryString("m")%>&Language=<%=language%>" class="qk"><%=word(548)%></a>   <a href="?step=2&m=<%=request.QueryString("m")%>&Language=<%=language%>" class="js"><%=word(549)%></a></td>
</tr>	
</form>
</table>

<%
else
%>
<div class="cart_none"><img src="<%=sitepath%>images/cart_search.gif" /><p><%=word(550)%></p><a href="<%=sitepath%>product/?Language=<%=language%>"><%=word(551)%> </a><div class="clear"></div></div>
</div> 
<%
End If
end If 
%>

<%
rs.close 
set rs = nothing 
'购物车结束
%>

<%
end if
'提交订单！填写购物详细内容 第二步
if request.QueryString("step")=2 then 

if session("MemName")="" and session("MemLogin")<>"Succeed" and ykgw=2 then
 writemsg1 2,""&word(552)&"","member/memberuserlogin/?m="&m2&"&Language="&Language&"","wrong"
else
%>
<div id="cart_div">
<div class="cart_quick">
<div class="cart_ico"><%=word(532)%></div>
<div class="cart_step">
<ul>
<li class="step1"><span class="active"><span><span><%=word(533)%></span></span></span></li>
<li class="step2"><span class="active"><span><span><%=word(534)%></span></span></span></li>
<li class="step3"><span><span><span><%=word(535)%></span></span></span></li>
</ul>
</div>
</div>
<table  border="0" cellpadding="0" cellspacing="0" class="productgwc">

<tr class="gwcwz">
<td width="9%" ><%=word(536)%></td>
<td width="30%"  ><%=word(537)%></td>
<td width="11%"  ><%=word(538)%></td>
<td width="9%" ><%=word(539)%></td>
<td width="15%" ><%=word(540)%></td>
<td width="17%" style="border-right:none"><%=word(541)%></td>
</tr>
<%
aryCart = Session("Cart")
i = 0
%>
<%
strTotalretailprice = 0
strTotalprice = 0
strcount = 0
For m = 0 to Ubound(aryCart,2)
%>
<tr class="gwcsj">
<td >
<%
if aryCart(i+6,m)=0 then
sql1="select * from product where p_sjs="&aryCart(i,m)&" and Lang='"&Language&"'"
set rs1=server.CreateObject("adodb.recordset")
rs1.open sql1,conn,1,1%>
<%if rs1("zmone")<>0 and rs1("zmtwo")=0 and rs1("zmthree")=0 then%>
<a href=<%=sitepath%>showproduct/?pone=<%=rs1("zmone")%>&proID=<%=rs1("id")%>>
<%end if%>
<%if rs1("zmone")<>0 and rs1("zmtwo")<>0 and rs1("zmthree")=0 then%>
<a href=<%=sitepath%>showproduct/?pone=<%=rs1("zmone")%>&ptwo=<%=rs1("zmtwo")%>&proID=<%=rs1("id")%>>
<%end if%>
<%if rs1("zmone")<>0 and rs1("zmtwo")<>0 and rs1("zmthree")<>0 then%>
<a href=<%=sitepath%>showproduct/?pone=<%=rs1("zmone")%>&ptwo=<%=rs1("zmtwo")%>&pthree=<%=rs1("zmthree")%>&proID=<%=rs1("id")%>>
<%end if%>
<%if aryCart(i+5,m)<>"" then%>
<img src="<%=sitepath%><%=aryCart(i+5,m)%>" width=80px height=40px/>
<%else%>
<img src="<%=sitepath%>images/demo.jpg" width=80px height=40px/>
<%end if%>
</a>

<%
set rs1=nothing
rs1.close

else


sql1="select * from product_stock where ID="&aryCart(i,m)&" and Lang='"&Language&"'"
set rs1=server.CreateObject("adodb.recordset")
rs1.open sql1,conn,1,1

sql2="select * from product where ID="&rs1("p_id")&" and Lang='"&Language&"'"
set rs2=server.CreateObject("adodb.recordset")
rs2.open sql2,conn,1,1

%>
<%if rs2("zmone")<>0 and rs2("zmtwo")=0 and rs2("zmthree")=0 then%>
<a href=<%=sitepath%>showproduct/?pone=<%=rs2("zmone")%>&proID=<%=rs2("id") %>>
<%end if%>
<%if rs2("zmone")<>0 and rs2("zmtwo")<>0 and rs2("zmthree")=0 then%>
<a href=<%=sitepath%>showproduct/?pone=<%=rs1("zmone")%>&ptwo=<%=rs1("zmtwo")%>&proID=<%=rs2("id") %>>
<%end if%>
<%if rs2("zmone")<>0 and rs2("zmtwo")<>0 and rs2("zmthree")<>0 then%>
<a href=<%=sitepath%>showproduct/?pone=<%=rs2("zmone")%>&ptwo=<%=rs2("zmtwo")%>&pthree=<%=rs2("zmthree")%>&proID=<%=rs2("id") %>>
<%end if%>
<%if aryCart(i+5,m)<>"" then%>
<img src="<%=sitepath%><%=aryCart(i+5,m)%>" width=80px height=40px/>
<%else%>
<img src="<%=sitepath%>images/demo.jpg" width=80px height=40px/>
<%end if%>
</a>

<%
set rs2=nothing
rs2.close

set rs1=nothing
rs1.close
%>
<%end if%></td> 
<td>
<%

if aryCart(i+6,m)=0 then
sql1="select * from product where p_sjs="&aryCart(i,m)&" and Lang='"&Language&"'"
set rs1=server.CreateObject("adodb.recordset")
rs1.open sql1,conn,1,1%>
<%if rs1("zmone")<>0 and rs1("zmtwo")=0 and rs1("zmthree")=0 then%>
<a href=<%=sitepath%>showproduct/?pone=<%=rs1("zmone")%>&proID=<%=rs1("id")%>>
<%end if%>
<%if rs1("zmone")<>0 and rs1("zmtwo")<>0 and rs1("zmthree")=0 then%>
<a href=<%=sitepath%>showproduct/?pone=<%=rs1("zmone")%>&ptwo=<%=rs1("zmtwo")%>&proID=<%=rs1("id")%>>
<%end if%>
<%if rs1("zmone")<>0 and rs1("zmtwo")<>0 and rs1("zmthree")<>0 then%>
<a href=<%=sitepath%>showproduct/?pone=<%=rs1("zmone")%>&ptwo=<%=rs1("zmtwo")%>&pthree=<%=rs1("zmthree")%>&proID=<%=rs1("id")%>>
<%
end if
set rs1=nothing
rs1.close

%>
<%=aryCart(i+2,m) %></a>
<%else%>

<%




sql1="select * from product_stock where ID="&aryCart(i,m)&" and Lang='"&Language&"'"
set rs1=server.CreateObject("adodb.recordset")
rs1.open sql1,conn,1,1

sql2="select * from product where ID="&rs1("p_id")&" and Lang='"&Language&"'"
set rs2=server.CreateObject("adodb.recordset")
rs2.open sql2,conn,1,1

%>
<%if rs2("zmone")<>0 and rs2("zmtwo")=0 and rs2("zmthree")=0 then%>
<a href=<%=sitepath%>showproduct/?pone=<%=rs2("zmone")%>&proID=<%=rs2("id") %>>
<%end if%>
<%if rs2("zmone")<>0 and rs2("zmtwo")<>0 and rs2("zmthree")=0 then%>
<a href=<%=sitepath%>showproduct/?pone=<%=rs1("zmone")%>&ptwo=<%=rs1("zmtwo")%>&proID=<%=rs2("id") %>>
<%end if%>
<%if rs2("zmone")<>0 and rs2("zmtwo")<>0 and rs2("zmthree")<>0 then%>
<a href=<%=sitepath%>showproduct/?pone=<%=rs2("zmone")%>&ptwo=<%=rs2("zmtwo")%>&pthree=<%=rs2("zmthree")%>&proID=<%=rs2("id") %>>
<%end if%>


<%=aryCart(i+2,m) %></a>

<%
' 购物车显示规格
sql3="select * from product_stock where ID="&aryCart(i,m)&" and Lang='"&Language&"'"
set rs3=server.CreateObject("adodb.recordset")
rs3.open sql3,conn,1,1

response.write("<br />"&word(543)&""&rs3("s_value")&" <br />"&word(544)&""&rs3("p_stock")&" ")

set rs3=nothing
rs3.close
%>




<%
end if
set rs2=nothing
rs2.close

set rs1=nothing
rs1.close
%>

</td> 	
<td ><%=formatnumber2(aryCart(i+3,m)) %> <%=hbwz%></td> 
<td ><%=formatnumber2(aryCart(i+4,m)) %> <%=hbwz%></td> 
<td ><%=aryCart(i+1,m)%></td> 
<td style="border-right:none"><%=hbfh%><%=formatNumber2(aryCart(i+1,m)*aryCart(i+4,m)) %></td>
</tr>
<%
strTotalretailprice=strTotalretailprice+aryCart(i+1,m)*aryCart(i+3,m)
strTotalprice=strTotalprice+aryCart(i+1,m)*aryCart(i+4,m)
strcount=strcount+aryCart(i+1,m)
Next
%>
</tr>
<tr class="gwcan">
<td colspan="7" style="border:0" >
<span><%=word(553)%><%=strcount%><%=word(554)%>  <%=word(555)%><label id="allpay"><%=hbfh%><%=formatnumber2(strTotalprice)%></label> <%=hbwz%></span> 
<a href="?step=1&m=<%=request.QueryString("m")%>&Language=<%=Language%>" class="fh"><%=word(556)%></a>
<%
session("auuont")=formatnumber2(strTotalprice)

%>
</td>
</tr>	

</table>

<form action="<%=sitepath%>inc/ProductBuy/?MemberID=<%=mMemID%>&m=<%=request.QueryString("m")%>&Language=<%=Language%>" method="post" class="checkform">
<table class="tjdd">
<tr>
<td class="zkd" ><%=word(557)%></td>
<td class="ckd" ><input name="OrderName" type="text"  id="OrderName" value="<%=word(558)%>"  class="inputxt"/></td>
<td class="ykd"><div class="Validform_checktip"><%=word(559)%></div></td>
</tr>
<tr>
<td class="zkd"><%=word(560)%></td>
<td class="ckd"><textarea  name="Remark"  id="Remark" ></textarea></td>
<td class="ykd"><div class="Validform_checktip"><%=word(561)%></div></td>
</tr>
<tr>
<td class="zkd"><%=word(562)%></td>
<td class="ckd" ><input  name="RealName" type="text" id="RealName" value="<%=mRealName%>"  class="inputxt" datatype="*2-6" nullmsg="<%=word(563)%>" errormsg="<%=word(564)%>" /><span class="fh">*</span></td>
<td class="ykd"><div class="Validform_checktip"><%=word(565)%></div></td>
</tr>
<tr>
<td class="zkd"><%=word(566)%></td>
<td class="ckd" >
<input  name="Sex" style="border:none" type="radio" value="<%=word(567)%>" <%if mSex=""&word(567)&"" then response.write ("checked")%> />
<%=word(567)%>
<input type="radio" name="Sex" style="border:none"  value="<%=word(568)%>" <%if mSex=""&word(568)&"" then response.write ("checked")%> />
<%=word(568)%>
</td>
<td class="ykd"><div class="Validform_checktip"></div></td>
</tr>
<tr>
<td class="zkd"><%=word(569)%></td>
<td class="ckd" ><input class="inputxt" name="Company" type="text" id="Company" value="<%=mCompany%>"  datatype="*" nullmsg="<%=word(570)%>" /><span class="fh">*</span></td>
<td class="ykd"><div class="Validform_checktip"><%=word(571)%></div></td>
</tr>
<tr>
<td class="zkd"><%=word(572)%></td>
<td class="ckd" ><input class="inputxt" name="Address" type="text" id="Address" value="<%=mAddress%>"  datatype="*" nullmsg="<%=word(573)%>"  /><span class="fh">*</span></td>
<td class="ykd"><div class="Validform_checktip"><%=word(574)%></div></td>
</tr>
<tr>
<td class="zkd"><%=word(575)%></td>
<td class="ckd" ><input class="inputxt" name="ZipCode" type="text" id="ZipCode" value="<%=mZipCode%>"  datatype="p" nullmsg="<%=word(576)%>" errormsg="<%=word(577)%>" /><span class="fh">*</span></td>
<td class="ykd"><div class="Validform_checktip"><%=word(578)%></div></td>
</tr>
<tr>
<td class="zkd"><%=word(579)%></td>
<td class="ckd" ><input class="inputxt" name="Fax" type="text" id="Fax" value="<%=mFax%>"  /></td>
<td class="ykd"><div class="Validform_checktip"></div></td>
</tr>
<tr>
<td class="zkd"><%=word(580)%></td>
<td class="ckd" ><input class="inputxt" name="Mobile" type="text" id="Mobile" value="<%=mMobile%>" datatype="m" errormsg="<%=word(581)%>" nullmsg="<%=word(582)%>"/><span class="fh">*</span></td>
<td class="ykd"><div class="Validform_checktip"><%=word(583)%></div></td>
</tr>

<tr>
<td class="zkd"><%=word(584)%></td>
<td class="ckd" ><input class="inputxt" name="Email" type="text" id="Email" value="<%=mEmail%>" datatype="e" errormsg="<%=word(585)%>" nullmsg="<%=word(586)%>"/><span class="fh">*</span></td>
<td class="ykd"><div class="Validform_checktip"><%=word(587)%></div></td>
</tr>
<tr>
<td class="zkd"><%=word(588)%></td>
<td class="ckd" >

<%
'配送方式
set rs=server.createobject("adodb.recordset")
sql="select * from shoppingshipment where lang='"&Language&"'"
rs.open sql,conn,1,1
if not rs.eof Then 
do while not rs.eof
%>
<input type="radio" style="border:none" name="shipment" id="shipment" datatype="*" nullmsg="<%=word(589)%>" value="<%=Trim(rs("ID"))%>|<%=formatnumber(Trim(rs("shipmentMoney")),2)%>|<%=Trim(rs("shipmentName"))%>" />
<%=Trim(rs("shipmentName"))%>(<%=shoppingcurrency%><%=Trim(rs("shipmentMoney"))%><%=fhwz%>)<br/><br/>

<span style="padding-left:20px"><%=Trim(rs("content"))%></span><br/><br/>

<%
rs.movenext
loop
end if
rs.close 
set rs = nothing 
%>

</td>
<td class="ykd"><div class="Validform_checktip"><%=word(590)%></div></td>
</tr>
<tr>
<td class="zkd"><%=word(591)%></td>
<td class="ckd" >

<select name="zffs" id="zffs" class="inputxt" datatype="*" nullmsg="<%=word(592)%>">
                            <option value=""><%=word(593)%></option>
                        	<option value=1><%=word(594)%></option>
                        	<option value=2><%=word(595)%></option>
                        </select><span class="fh">*</span>

</td>
<td class="ykd"><div class="Validform_checktip"><%=word(596)%></div></td>
</tr>
<tr>
<td class="zkd"><%=word(597)%></td>
<td class="ckd" >
<input class="inputxt" name="captchacode" type="text" id="captchacode"  style="width:65px;" datatype="*"  nullmsg="<%=word(598)%>"  ajaxurl="<%=sitepath%>ajax/yzm.asp?language=<%=language%>"/> 
<img id="imgCaptcha" src="<%=SITEPATH%>inc/captcha.asp" /> <a href="###" onClick="RefreshImage('imgCaptcha')"><img src="<%=sitepath%>images/reload.png" width="20" height="20" /></a>
</td>
<td class="ykd"><div class="Validform_checktip"></div></td>
</tr>
<tr>
<td colspan="3"  class="last">
  <input name="Submit" type="submit" class="tjan" value="<%=word(599)%> "  />
  <input name="Reset" type="reset" class="czan" value=" <%=word(600)%> " /></td>
</tr>

</table>
</form>
</div> 
<%
end if
end if
%>
<!--第二步结束-->

</div>
</div>


<!--#include file="../bottom.asp"-->
</body>
</html>
