<!--#include file="Inc/Function.asp"-->
<%
sltdz=request.QueryString("sltdz")
session("sltdz")=sltdz
call deletefile1(sltdz)

%>

<script LANGUAGE="JavaScript">
top.window.opener = top;
top.window.open('','_self','');
top.window.close();

 </script>