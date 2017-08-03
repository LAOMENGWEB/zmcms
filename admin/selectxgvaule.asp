<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="Inc/Function.asp"-->
<%
if Instr(session("AdminPurview"),"|36,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</script>"
  response.end
end if
id=request.QueryString("id")
s_id=request.QueryString("s_id")
mode=request.QueryString("mode")

If request.Form("submit") = "修改" Then
s_value=Trim(request.form("s_value"))
s_order=trim(request.form("s_order"))
Set rs = server.CreateObject("adodb.recordset")
rs.open "Select * From SelectValue Where id=" & id & "",conn,1,3

rs("s_value")=s_value
rs("s_order")=s_order
rs("lang")=lang
rs.update
rs.close
set rs=nothing 
call addlog("修改属性值")

response.Write "<script language=javascript>alert(""恭喜您修改成功"");window.location.href=""selectglvaule.asp?ID="&s_id&"&mode="&mode&"""</script>"


end if


Set rs3 = server.CreateObject("adodb.recordset")
rs3.open "Select * From Selectvalue Where id=" & id & " and lang='"&lang&"'",conn,1,3

If rs3.EOF Then
    response.Write "<script language=javascript>window.location.href=""shuxing.asp""</script>"
    response.End
End if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>附属属性类别增加</title>
<script language="JavaScript" src="js/jquery-1.7.2.min.js"></script>
 <script type="text/javascript" src="../js/Validform_v5.3.2_min.js"></script>
 <link href="css/style.css" rel="stylesheet" type="text/css">
  <script type="text/javascript">
$(function(){
	
$(".checkform").Validform();  //就这一行代码！;
	
})
</script>

<style type="text/css">
.table1{ font-size: 14px; background: #F9f9f9;border:none;}
.table1 td{padding:10px; font-family:"微软雅黑";color:#ccc; font-weight:bold;}
.table1 td a{color:#ffffff}
</style>
<!--#include file="hqyy.asp"-->
</head>

<body>

<form method="POST" name="myform" class="checkform">


<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center"><%if mode=0 then%>修改筛选属性值名称<%end if%>
     <%if mode=1 then%>修改规格属性值名称<%end if%>
     <%if mode=2 then%>修改商品属性值名称<%end if%> <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>



<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="shuxing.asp" target="main">返回菜单</a></div></td>
 

 </tr>
</table>


<table  width="99%"  border="0" cellpadding="6" cellspacing="0" class="dtable4" align="center">
<tr > 
 <td width="11%"   >属性值名称：</td>
  <td width="89%">
<input  name="s_value" type="text" id="s_value" value="<%=rs3("s_value")%>"  size="20" datatype="*" nullmsg="属性值名称不能为空" /> <font color="#FF0000">*</font></td>
</tr>

<tr > 
 <td  >排序：</td>
  <td>
<input name="s_order" type="text" id="s_order"   value="<%=rs3("s_order")%>" style="width:40px"></td>
</tr>
<tr > 
<td colspan="2"> <input class="bnt" type="submit" name="Submit" value="修改">
</td>
</tr>       
</table>
</form >
<%
set rs3=nothing
rs3.close
%>
</body>
</html>

