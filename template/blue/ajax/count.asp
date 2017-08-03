<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../../../Inc/sub.asp" --> 
<%
aryCart = Session("Cart")
if arycart<>"" then
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
session("totalcount")=strcount
else
session("totalcount")=0
end if
response.write(""&session("totalcount")&"")
%>	
