<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../../../Inc/sub.asp" --> 

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>获取静态页下面的价格</title>
</head>

<body>
<%
proid=request.QueryString("proid")

Set rspro=Server.CreateObject("ADODB.RecordSet") 
sql="select * from product where id="&proid&""
rspro.Open sql,conn,1,1
If Trim(RsPro("P_SpecificationName")) <> "" And Not isnull(Trim(RsPro("P_SpecificationName"))) Then 
if session("MemName")="" or session("GroupID")="" or session("GroupID")=999177 or session("MemLogin")<>"Succeed"   then
Set rsprice = conn.execute("select top 1 p_price from [Product_Stock] where p_id = "& proid &" and Lang='"&Language&"' order by p_price asc")
If Not rsprice.eof Then 
Lowprice = formatnumber2(rspro("hyPrice") + rsprice(0))
End If 
rsprice.close : Set rs = Nothing
Set rsprice = conn.execute("select top 1 p_price from [Product_Stock] where p_id = "& proid &" and Lang='"&Language&"' order by p_price desc")
If Not rsprice.eof Then 
Highprice =formatnumber2( rspro("hyPrice") + rsprice(0))
End If 
rsprice.close : Set rs = Nothing
If CDbl(Highprice) = CDbl(Lowprice) Then 
response.write(lowprice)
Else
response.write(""&lowprice&"-"&Highprice&"")
End If 															
else
set rsjg=server.CreateObject("adodb.recordset")
sqljg="select * from hygroup where groupid='"&session("GroupID")&"' and Lang='"&Language&"'"
rsjg.open sqljg,conn,1,1
set rsjg1=server.CreateObject("adodb.recordset")
sqljg1="select * from product_price where grade_id="&rsjg("GroupLevel")&" and product_id="&proid&" and Lang='"&Language&"'"
rsjg1.open sqljg1,conn,1,1
Set rsprice = conn.execute("select top 1 p_price from [Product_Stock] where p_id = "& proid &" and Lang='"&Language&"' order by p_price asc")
If Not rsprice.eof Then 
Lowprice = formatnumber2(rsjg1("Price") + rsprice(0))
End If 
rsprice.close : Set rs = Nothing
Set rsprice = conn.execute("select top 1 p_price from [Product_Stock] where p_id = "& proid &" and Lang='"&Language&"' order by p_price desc")
If Not rsprice.eof Then 
Highprice =formatnumber2( rsjg1("Price") + rsprice(0) )
End If 
rsprice.close : Set rs = Nothing
If CDbl(Highprice) = CDbl(Lowprice) Then 
response.write(lowprice)
Else
response.write(""&lowprice&"-"&Highprice&"")
End If 

end if						
else
if session("MemName")="" or session("GroupID")="" or session("GroupID")=999177 or session("MemLogin")<>"Succeed"   then
response.write(formatnumber2(rspro("hyPrice")))
else
set rsjg=server.CreateObject("adodb.recordset")
sqljg="select * from hygroup where groupid='"&session("GroupID")&"' and Lang='"&Language&"'"
rsjg.open sqljg,conn,1,1
set rsjg1=server.CreateObject("adodb.recordset")
sqljg1="select * from product_price where grade_id="&rsjg("GroupLevel")&" and product_id="&proid&" and Lang='"&Language&"'"
rsjg1.open sqljg1,conn,1,1
response.write(formatnumber2(rsjg1("Price")))

end if
end if

set rspro=nothing
rspro.close

%>

</body>
</html>
