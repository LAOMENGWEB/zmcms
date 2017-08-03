<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="Inc/Function.asp"-->
<%
if Instr(session("AdminPurview"),"|1000004,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</script>"
  response.end
end if
'========判断是否具有管理权限
%>
<%
If request.Form("submit") = "添加" Then
province=Trim(request.form("province"))
xs=trim(request.form("xs"))
m=trim(request.form("m"))
px=trim(request.form("px"))
	set rs = server.CreateObject ("adodb.recordset")
	sql="select * from province Where province='"&province&"'"
	rs.open sql,conn,1,3
	if rs.eof and rs.bof then
		rs.AddNew 
		rs("province")=province
		rs("xs")=xs	
		rs("m")=m
		rs("px")=px
		rs("lang")=lang
		rs.update
		rs.close
	set rs=nothing
	call addlog("添加营销区域")
		response.Write "<script language=javascript>alert(""恭喜您添加成功"");window.location.href=""sf.asp""</script>"
	Else
		response.Write "<script language=javascript>alert(""该营销区域已经存在"");window.history.back(-1);</script>"
response.End
	End if
	
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


<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">添加区域  <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>


<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="chajian.asp" target="main">返回菜单</a></div><div class="dmeun"><a href="sf.asp" target="main">区域管理</a></div></td>
 

 </tr>
</table>
<form  name="myform" class="checkform" method="post">
<table width="99%"  border="0" cellpadding="0" cellspacing="0" class="dtable4" align="center">

<tr> 
<td width="12%" >区域名称:</td>
<td width="88%">
  <input name="province" type="text" id="province" size="40" maxlength="50" datatype="*" nullmsg="区域不能为空"></td>
</tr>
<tr>
  <td >高亮值:</td><td>
    <input name="m" type="text" id="m" value="0" size="40" maxlength="50"></td>
  </tr>
   <tr>
       <td>是否显示</td>
       <td><input type="radio" class="none"  name="xs" value="1" checked> 
是
 <input type="radio" class="none"  name="xs" value="0" >
否	</td>
     </tr>
<tr>
  <td >排序:</td><td>
    <input name="px" type="text" id="px" value="0" size="8" maxlength="5" >  </td>
  </tr>
<tr> 
<td  colspan="2"><input name="Submit" type="submit" class="bnt" value="添加"></td>
</tr>

</table>
</form>
</body>
</html>
