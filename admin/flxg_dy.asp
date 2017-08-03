<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<%
ID=request("ID")
ClassName=request("ClassName")
px=request("px")
m=Request("m")
if request("Submit")="修改" then
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql="select * from mnclass where id="&ID
	rs.open sql,conn,1,3
    rs("m")=m
	rs("ClassName")=ClassName
    rs("px")=px
    rs("lang")=lang
	rs.update
	rs.close
	Set rs=nothing
	call addlog("修改"&danye&"类别")
	response.Write "<script language=javascript>alert(""类别修改成功"");window.location.href=""fl_dy.asp""</script>"
    response.End
end if
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
<%
Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from mnclass where id = " & ID&" and lang='"&lang&"'"
rs.open sql,conn,1,1
If rs.EOF Then
    response.Write "<script language=javascript>window.location.href=""fl_dy.asp""</script>"
    response.End
End if
%>
<form id="zmcms" name="zmcms" method="post" class="checkform">

<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">修改<%=danye%>分类  <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>
<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="danye.asp" target="main">返回菜单</a></div><div class="dmeun"><a href="fl_dy.asp" target="main">分类管理</a></div></td>
 

 </tr>
</table>
  <table width="99%"  border="0" cellpadding="0" cellspacing="0" align="center" class="dtable4">
   
    <tr>
      <td width="9%">类别名称:</td>
     <td width="91%"> <input name="classname" type="text" id="classname" value="<%=rs("classname")%>" datatype="*" nullmsg="类别名称不能为空"/>
  
      </td>
    </tr>
            <tr>
      <td>高亮值:</td><td>
      <input name="m" type="text" id="m" value="<%=rs("m")%>"/>
      高亮值 必须 和导航菜单的相关栏目 高亮值相同。  
      </td>
     
    </tr> 
	 <tr>
      <td>排序:</td>
     <td>
      <input name="px" type="text" id="px" value="<%=rs("px")%>" size="5" />

      </td>
    </tr>
	 <tr>
      <td colspan="2">
     <label>
      <input type="submit" name="Submit" value="修改" class="bnt"/>
      </label>
      </td>
    </tr>
  </table>
</form>
<%
rs.close
set rs=nothing
%>
</body>
</html>
