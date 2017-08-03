<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../../../Inc/sub.asp" --> 
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>获取属性信息</title>
</head>

<body>
<%
Select Case request("action")
Case "getspinfo" : getspinfo()
End Select

Sub getspinfo()

name = safereplace(request("name"))
id = checknum(request("id"),0)
price = checknum(request("price"),0)
Set rs = conn.execute("select p_stock,p_price,s_value,id,p_no from [Product_Stock]  where p_id = "& id &" and s_value='"& name &"' and Lang='"&Language&"' ")
If Not rs.eof Then 
	%>
||||<%=word(623)%><%=hbfh%><%=formatnumber2((price+rs(1)))%> 
<%
Else
response.write("0||||")
End If 

sql21="update product set s_value='"&name&"' where id in ("&id&")"
conn.Execute ( sql21)
sql22="update product set p_id='"&rs(3)&"' where id in ("&id&")"
conn.Execute ( sql22)
sql23="update product set p_price='"&rs(1)&"' where id in ("&id&")"
conn.Execute ( sql23)
sql24="update product set p_no='"&rs(4)&"' where id in ("&id&")"
conn.Execute ( sql24)
rs.close : Set rs = Nothing
End Sub 
%>

</body>
</html>

