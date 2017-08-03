<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!-- #include file="admin.asp" -->
<!-- #include file="Inc/function.asp" -->
<%
if Instr(session("AdminPurview"),"|91,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</script>"
  response.end
end if

Set rs = server.CreateObject("adodb.recordset")
sql = "select * from waplogo where lang='"&lang&"' and skins='"&skins&"'  "
rs.Open sql, conn, 1, 3
If Not(rs.Bof and rs.Eof) Then 
x=1
else
x=2
end if
rs.close
set rs=nothing

'========判断是否具有管理权限
If request.Form("submit") = "保存" Then

waplogourl=request("waplogourl")
waplogowidth=request("waplogowidth")
waplogoheight=request("waplogoheight")
Set rs = server.CreateObject("adodb.recordset")
if x=1 then
sql = "select * from waplogo where lang='"&lang&"' and skins='"&skins&"'   "
else
sql = "select * from waplogo  "
end if
rs.Open sql, conn, 1, 3

if x=2 then
rs.addnew
end if

rs("waplogourl")=waplogourl
rs("skins")=skins
rs("lang")=lang
rs("waplogowidth")=waplogowidth
rs("waplogoheight")=waplogoheight

rs.update
    rs.Close
    Set rs = Nothing
   
	call addlog("设置手机端LOGO成功!")
  response.Write "<script language=javascript>alert(""设置成功"");window.location.href=""waplogoset.asp""</script>"
    response.End
End If
%>


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>zmcms_管理页面</title>

<script language="javascript" src="../js/Admin.js"></script>
<script type="text/javascript" src="js/jquery-1.7.2.min.js"></script>
<script type="text/javascript" src="../js/Validform_v5.3.2_min.js"></script>
<script type="text/javascript" src="../javascript/colorbox/jquery.colorbox-min.js"></script>
<link href="../javascript/colorbox/colorbox.css" rel="stylesheet" media="screen">
<link rel="stylesheet" href="../zmbjq/themes/default/default.css" />
<link rel="stylesheet" href="../zmbjq/plugins/code/prettify.css" />
<script charset="utf-8" src="../zmbjq/kindeditor-all.js"></script>
<script charset="utf-8" src="../zmbjq/lang/zh_CN.js"></script>
<script charset="utf-8" src="../zmbjq/plugins/code/prettify.js"></script>
<script language=javascript>
$(document).ready(function(){
$(".allfile").colorbox({iframe: true, innerWidth: 1000,innerHeight: 500});
$(".sfile").colorbox({iframe: true, innerWidth: 400,innerHeight: 160});
});
</script>

<script type="text/javascript">
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

<link href="css/style.css" rel="stylesheet" type="text/css">
 <script type="text/javascript">
$(function(){
	
$(".checkform").Validform();  //就这一行代码！;
	
})
</script>
<style type="text/css">
table{ font-size: 14px;  border-right:none}
table td{padding:10px; font-family:"微软雅黑";color:#ccc; font-weight:bold;}
table td a{color:#ffffff}
</style>
<!--#include file="hqyy.asp"-->
</head>


<body>

<form name="editForm" method="POST" class="checkform"  >
<%
Set rs = server.CreateObject("adodb.recordset")
sql = "select * from waplogo where lang='"&lang&"' and skins='"&skins&"'"
rs.Open sql, conn, 1, 1
%>		
<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" style=" background: #F9f9f9;border:none; ">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">手机端LOGO设置  <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>

<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="shouji.asp" target="main">返回菜单</a></div></td>
 

 </tr>
</table>

<table width="99%" border="0" align="center" cellpadding="0" cellspacing="0"  class="dtable4">

   
 
  
  


      <tr >
<td>LOGO图片：</td>
			    <td colspan="2">
				   <input name="waplogourl"  type="TEXT" id="weepic"  value="<%=rs("waplogourl")%>" size="50"  /> 
				  <br />
				   <br />
			      <a class='sfile' href="Upload_All.Asp?add=one" title="支持上传格式(*.swf;*.gif;*.bmp)"><div style="width:100px;padding:10px; background:#28b576;color:#ffffff" align="center">手机LOGO上传</div></a>
		        <div id="mysmaimg">
				<%if x=1 then%>
<img src="../<%=rs("waplogourl")%>" id="sc"/><br />
<a href="delslt.asp?sltdz=<%=rs("waplogourl")%>" target="_blank" onClick="deltp()" id="lj">删除</a>
<%end if%></div></td>
    </tr>
		
<tr> 
<td>LOGO大小：</td>
<td colspan="2">
	  
          <input name="waplogowidth"type="text" class="inputtext" id="logowidth"  value="<%=rs("waplogowidth")%>" size="6" maxlength="30" />
                    logo宽×高
        <input name="waplogoheight"type="text" class="inputtext" id="logoheight"  value="<%=rs("waplogoheight")%>" size="6" maxlength="30" /></td>
    </tr>
     
      <tr>
	  
    <td colspan="2">
      <input type="submit" name="Submit" value="保存" class="bnt" />   </td>
  </tr>
</table>
</form>

</body>
</html>
<%
rs.nothing
set rs=nothing
 conn.Close
    Set conn = Nothing
%>
