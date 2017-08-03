<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="Inc/Function.asp"-->
<%
If request.Form("submit") = "修改"  Then
id=request("id")
langname=Trim(request.form("langname"))
langbs=trim(request.form("langbs"))
huobi=trim(request.form("huobi"))
huobi1=trim(request.form("huobi1"))
SmallImage=trim(request.form("SmallImage"))
px=trim(request.form("px"))

set rs = server.CreateObject ("adodb.recordset")
	sql="select * from lang where id="&id
	rs.open sql,conn,1,3
rs("langname")=langname
		rs("langbs")=langbs	
	 rs("huobi")=huobi
	  rs("huobi1")=huobi1
	  rs("SmallImage")=SmallImage
	
	  rs("px")=px
		rs.update
	rs.close
	set rs=nothing
	call addlog("修改多语言")
	response.Write "<script language=javascript>alert(""恭喜您修改成功"");window.location.href=""lang.asp""</script>"
	end if
Set rs1= server.CreateObject("adodb.recordset")
sql2 = "select * from lang where id = "&request.QueryString("id")&""
rs1.Open sql2, conn, 1, 1
If rs1.EOF Then
    response.Write "<script language=javascript>alert(""不存在此数据！"");window.history.back();</script>"
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
<script type="text/javascript" src="../javascript/colorbox/jquery.colorbox-min.js"></script>
<link href="../javascript/colorbox/colorbox.css" rel="stylesheet" media="screen">
<script language=javascript>
$(document).ready(function(){
$(".allfile").colorbox({iframe: true, innerWidth: 1000,innerHeight: 500});
$(".sfile").colorbox({iframe: true, innerWidth: 400,innerHeight: 160});
});
</script>
<script language=javascript>
function deltp()
{

var parent=document.getElementById("mysmaimg");
var child=document.getElementById("sc");
var child1=document.getElementById("lj");
parent.removeChild(child);
parent.removeChild(child1);
document.getElementById("weepic").value="";
}

 </script>
</head>

<body>
<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">修改语言</div></td>
</tr>
</table>




<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="lang.asp" target="main">多语言管理</a></div></td>
 

 </tr>
</table>
<form  name="myform" method=post class="checkform">
<table width="99%"  border="0" cellpadding="0" cellspacing="0" class="dtable4" align="center">

<tr> 
<td width="12%" >语言名称:</td>
<td width="88%">
  <input name="langname" type="text" id="langname" size="40" maxlength="50" datatype="*" nullmsg="语言名称不能为空" value="<%=rs1("langname")%>"></td>
</tr>
<tr>
  <td >语言标识:</td><td>
    <input name="langbs" type="text" id="langbs" value="<%=rs1("langbs")%>" size="40" maxlength="50" datatype="*" nullmsg="语言标识不能为空"></td>
  </tr>



			    <td >语言图标：
		 </td>
<td colspan="2">  <input name="SmallImage"  type="hidden" id="weepic"  value="<%=rs1("SmallImage")%>" size="50"  />	<br/><br/>	<a class='sfile' href="Upload_All.Asp?add=one" title="支持上传格式(*.swf;*.gif;*.bmp)"><img src="Images/dtup.jpg"  border="0"/></a><br />
<div id="mysmaimg"><%if rs1("Smallimage")<>"" then%>
<img src="<%=sitepath%><%=rs1("SmallImage")%>" id="sc" class="wd"/><br />
<a href="delslt.asp?sltdz=<%=rs1("Smallimage")%>" target="_blank" onclick="deltp()" id="lj">删除</a>
<%end if%>
</div></td>
	      </tr>
  <tr>
  <td >货币符号:</td><td>
    <input name="huobi" type="text" id="langbs" value="<%=rs1("huobi")%>" size="40" maxlength="50" datatype="*" nullmsg="货币符号不能为空"></td>
  </tr>
  <tr>
  <td >货币文字:</td><td>
    <input name="huobi1" type="text" id="langbs" value="<%=rs1("huobi1")%>" size="40" maxlength="50" datatype="*" nullmsg="货币文字不能为空"></td>
  </tr>
   <tr>
  <td >排序:</td><td>
    <input name="px" type="text" id="px" value="<%=rs1("px")%>" size="40" maxlength="50" datatype="*" ></td>
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
