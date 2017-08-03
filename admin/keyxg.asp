<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="Inc/Function.asp"-->
<%
If request.Form("submit") = "修改"  Then
id=request("id")
Title=Trim(request.form("Title"))
Url=trim(request.form("Url"))
Num=zmcmsRequest(request.form("Num"))
YReplace=zmcmsRequest(request.form("Replace"))
set rs = server.CreateObject ("adodb.recordset")
	sql="select * from gjz where id="&id
	rs.open sql,conn,1,3
		rs("Title")=Title
		rs("Url")=Url	
		rs("Num")=Num
		rs("Replace")=YReplace
		rs("lang")=lang
		rs.update
	rs.close
	set rs=nothing
	call addlog("修改内链关键字")
	response.Write "<script language=javascript>alert(""恭喜您修改成功"");window.location.href=""key.asp""</script>"
	end if
Set rs1= server.CreateObject("adodb.recordset")
sql2 = "select * from gjz where id = "&request.QueryString("id")&" and lang='"&lang&"'"
rs1.Open sql2, conn, 1, 1
If rs1.EOF Then
    response.Write "<script language=javascript>window.location.href=""key.asp""</script>"
    response.End
End If

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>zmcms_管理页面</title>
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



<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">修改内链关键字  <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>

<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="chajian.asp" target="main">返回菜单</a></div><div class="dmeun"><a href="key.asp" target="main">内链关键字管理</a></div></td>
 

 </tr>
</table>
<form  name="myform" method=post class="checkform">
<table width="99%"  border="0" cellpadding="0" cellspacing="0" class="dtable4" align="center">

<tr> 
<td width="8%" >关键字: </td>
<td width="92%">
  <input name="title" type="text" id="title" value="<%=rs1("title")%>" size="40" maxlength="50"></td>
</tr>
<tr>
  <td >链接地址:</td><td>
    <input name="Url" type="text" id="Url" value="<%=rs1("url")%>" size="40" maxlength="50"></td>
  </tr>
<tr>
  <td >优先级:</td><td>
    <input name="Num" type="text" id="Num" value="<%=rs1("NUM")%>" size="8" maxlength="5" >   
     数字越大越优先</td>
  </tr>
<tr>
  <td >替换次数:</td><td>
    <input name="Replace" type="text" id="Replace" value="<%=rs1("Replace")%>" size="8" maxlength="5">   
    一页中此关键字替换的次数,0表示不限</td>
  </tr>
<tr> 
<td  colspan="2"><input name="Submit" type="submit" class="bnt" value="修改"></td>
</tr>

</table>
</form>
<%

rs.close
set rs=nothing
conn.close
	set conn=nothing
%>
</body>
</html>
