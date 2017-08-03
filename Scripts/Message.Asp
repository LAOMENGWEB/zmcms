<!--#include file="../inc/sub.asp"-->

<%
BContent = Trim(request("contenthide"))
BPhone = Trim(request("phonehide"))
BAddr = Trim(request("addrhide"))
BEmail = Trim(request("emailhide"))
set rs = server.createobject("adodb.recordset")
sql="select * from jszx"
rs.open sql,conn,1,3
rs.addnew
rs("BizContent") = BContent
rs("BizPhone") = BPhone
rs("BizAddr") = BAddr
rs("BizEMail") = BEmail
rs("BizDate") = now()
rs.update
rs.close
set rs = Nothing

%>