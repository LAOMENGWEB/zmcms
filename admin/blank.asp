<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!-- #include file="admin.asp" -->
<!-- #include file="Inc/function.asp" -->

<%
if Instr(session("AdminPurview"),"|515,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</script>"
  response.end
end if
'========判断是否具有管理权限
%>




<%  



If request.Form("submit") = "保存" Then
blank = request("blank")



Set rs = server.CreateObject("adodb.recordset")
sql = "select * from wzset where lang='"&lang&"' "
rs.Open sql, conn, 1, 3
rs("blank")= blank 
rs("lang")=lang
rs.update
    rs.Close
    Set rs = Nothing
   
	call addlog("修改银行转账信息")
  response.Write "<script language=javascript>alert(""修改银行转账信息"");window.location.href=""blank.asp""</script>"
    response.End
End If
%>


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>zmcms_管理页面</title>

<link rel="stylesheet" href="../zmbjq/themes/default/default.css" />
<link rel="stylesheet" href="../zmbjq/plugins/code/prettify.css" />
<script charset="utf-8" src="../zmbjq/kindeditor-all.js"></script>
<script charset="utf-8" src="../zmbjq/lang/zh_CN.js"></script>
<script charset="utf-8" src="../zmbjq/plugins/code/prettify.js"></script>


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

<form name="editForm" method="POST" action="blank.asp"  class="checkform">
<%
Set rs = server.CreateObject("adodb.recordset")
sql = "select * from wzset where lang='"&lang&"'"
rs.Open sql, conn, 1, 1
If rs.EOF Then
response.Write "<script language=javascript>window.location.href=""blank.asp""</script>"
response.End
End If


%>
		

<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center"> 银行转账设置 <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>


<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td >
 <div class="dmeun"><a href="chajian.asp" target="main">返回菜单</a></div>
 
 
 </td>
 

 </tr>
</table>

<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td >

 <div class="dmeun" style=" background:#eeeeee; color:#999999"><a href="alipay.asp" target="main" style=" color:#999999">支付宝设置</a></div>
 <div class="dmeun" ><a href="blank.asp" target="main" >银行转账设置</a></div>
 
 
 </td>
 

 </tr>
</table>
<table width="99%" border="0" align="center" cellpadding="0" cellspacing="0"  class="dtable4">
  
  <script>
 var editor;
 KindEditor.ready(function(K) {
 editor = K.create('textarea[name="blank"]', {
 filterMode : false, 
 uploadJson : '../zmbjq/asp/upload_json.asp',
 fileManagerJson : '../zmbjq/asp/file_manager_json.asp',
 allowFileManager : true,
 width: '700px',
 hight: '300px',
 afterBlur: function(){this.sync();}
 });

 });
 </script>
     <tr>
    <td >
      <label>
     <textarea name="blank" id="blank" style="width:700px;height:300px;visibility:hidden;"><%=rs("blank")%></textarea>
      </label></td>
  </tr>
   
 
    

      <tr>
    <td><label>
      <input type="submit" name="Submit" value="保存" class="bnt" />
    </label></td>
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
