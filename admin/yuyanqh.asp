<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="Inc/Function.asp"-->
<%
call initCodecs 

htmlurl=base64Encode(Request("htmlurl"))

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>zmcms_管理页面</title>
</head>

<body>


		语言切换：<select class="wbk" name="zmClassID" id="zmClassID" onChange="javascript:window.open(this.options[this.selectedIndex].value,'main')">
    <%
   Dim Sqlp,Rsp,TempStr
   Sqlp ="select * from lang  order by id asc"   
   Set Rsp=server.CreateObject("adodb.recordset")   
   rsp.open sqlp,conn,1,1 

      Do while not Rsp.Eof   
	      %>
     <option value="yuyan.asp?id=<%=rsp("id")%>&url=<%=htmlurl%>"
         <%
		 If lang=Rsp("langbs") then
				Response.Write(" selected" ) 
		 End if
         Response.Write(">" & Rsp("langname") & " </option>")  & VbCrLf
		 
      Rsp.Movenext   
      Loop   
 
	rsp.close
	set rsp=nothing
	%>
        </select>



</body>
</html>
