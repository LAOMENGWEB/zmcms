
<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../inc/conn.asp" -->

<%id=request.QueryString("id")
Set rsyxwl1=Server.CreateObject("ADODB.RecordSet") 
sql="select * from NetWork where id="&id
rsyxwl1.Open sql,conn,1,1
dtlx=rsyxwl1("dtlx")
%>

<%=dtlx%>

