<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!-- #include file="../../../inc/sub.asp" --> 
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>获取购物车按钮</title>
</head>

<body>
<%
set rs=server.CreateObject("adodb.recordset")
sql="select * from product where ID="&Trim(Request("ID"))&" and Lang='"&Language&"'"
rs.open sql,conn,1,1
s_value=rs("s_value")

set rs1=server.CreateObject("adodb.recordset")
sql1="select * from product_stock where s_value='"&s_value&"' and p_id="&Trim(Request("ID"))&" and Lang='"&Language&"'"
rs1.open sql1,conn,1,1

if not rs1.eof Then
'判断规格数量是否为空
if rs1("p_stock")<>0 then
response.write("<p class='msg' style=""font-size:12px;font-weight:normal;margin:0 0 7px ;""><i class='icon icon-right'></i><span style=""color:#ff0000"">"&word(525)&""&s_value&","&word(526)&""&rs1("p_stock")&","&word(527)&""&rs1("p_buys")&"</span></p><br /><img src="""&template&"Img/btn-cart.png"" id=""add_to_my_cart_img""  onClick=""add_2_my_cart();""/>")
else
response.write("<p class='msg' style=""font-size:12px;font-weight:normal;margin:0 0 7px;""><i class='icon icon-alert'></i><span style=""color:#ff0000"">"&word(528)&""&s_value&""&word(529)&"</span></p><br /><img src="""&template&"Img/btn-cart.png"" id=""add_to_my_cart_img""  />")
end if
else
response.write("<p class='msg' style=""font-size:12px;font-weight:normal;margin:0 0 7px;""><i class='icon icon-alert'></i><span style=""color:#ff0000"">"&word(530)&"</span></p><br /><img src="""&template&"Img/btn-cart.png"" id=""add_to_my_cart_img""  />")
end if
set rs1=nothing
rs1.close
set rs=nothing
rs.close
%>
</body>
</html>
