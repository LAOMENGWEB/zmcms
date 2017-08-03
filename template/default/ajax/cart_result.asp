<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../../../Inc/sub.asp" -->  
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>追梦工作室("www.zmxt.cn;技术支持:2685863740")</title>
<style type="text/css">
body{font-size:12px;}
</style>
</head>
<body>

<%
aryCart = Session("Cart")
i = 0
strTotalretailprice = 0
strTotalprice = 0
strcount = 0
totalcount=0
For m = 0 to Ubound(aryCart,2)
strTotalretailprice=strTotalretailprice+aryCart(i+1,m)*aryCart(i+3,m)
strTotalprice=strTotalprice+aryCart(i+1,m)*aryCart(i+4,m)
strcount=strcount+1
totalcount=totalcount+aryCart(i+1,m)
Next
%>	
<%=word(521)%></br></br>

<%=word(522)%><font style="color:red; font-weight:bold"><%=strcount%></font> <%=word(523)%>    <%=word(524)%><%=hbfh%><font style="color:red; font-weight:bold"><%=formatNumber2(strTotalprice)%></font> <%=hbwz%>
	
</body>
</html>
