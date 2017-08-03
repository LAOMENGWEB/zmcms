<!--#include file="qqconnect.asp"-->
<%
Dim qc, url
SET qc = New QqConnet
    Session("State")=qc.MakeRandNum()
    url = qc.GetAuthorization_Code()
    Response.Redirect(url)
Set qc=Nothing
%>