<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="admin.asp"-->
<!--#include file="conn.asp"-->
<%
if Instr(session("AdminPurview"),"|10413,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</script>"
  response.end
end if
'========判断是否具有管理权限
%>
 <%
  session("title2")=""
session("content2")=""
session("name2")=""
  %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>追梦工作室邮件发送系统</title>
<script language="JavaScript" src="js/jquery-1.7.2.min.js"></script>
 <script type="text/javascript" src="../js/Validform_v5.3.2_min.js"></script>
 <link rel="stylesheet" href="../zmbjq/themes/default/default.css" />
<link rel="stylesheet" href="../zmbjq/plugins/code/prettify.css" />
<script charset="utf-8" src="../zmbjq/kindeditor-all.js"></script>
<script charset="utf-8" src="../zmbjq/lang/zh_CN.js"></script>
<script charset="utf-8" src="../zmbjq/plugins/code/prettify.js"></script>
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
</head>
<body>
<%select Case Request("l")%>
<%
case "somail"
%>
  <form id="myform" name="myform" method="post" action="sendemail.asp?l=pumailgo"  class="checkform">





<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">邮件发送</div></td>
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

 <div class="dmeun" ><a href="sendemail.asp?l=somail" target="main" >手动发送邮件</a></div>
 <div class="dmeun" style=" background:#eeeeee; color:#999999"><a href="sendmemberemail.asp?l=somail" target="main" style=" color:#999999">所有订阅用户</a></div>
 <div class="dmeun" style=" background:#eeeeee; color:#999999"><a href="sendhyemail.asp?l=somail" target="main" style=" color:#999999">会员发送</a></div>
 
 </td>
 

 </tr>
</table>
		












  <table width="99%" border="0" align="center" cellpadding="0" cellspacing="0" class="dtable4">
  
    <tr>
      <td>收件人邮箱：</td> <td><input name="lmail" type="text" id="lmail" style="border:1px #ddd solid" value="<%
= Request("selectemail") %>" size="100" datatype="*" nullmsg="收件人不能为空"/> 如果发给多人请用,隔开邮件地址。</td>
        
        
    </tr>
    <tr>
      <td >邮件标题 ：</td>
       <td> <input name="title" type="text" id="title" value="<%=request("title")%>" size="40" datatype="*" nullmsg="邮件主题不能为空"/>    </td>
    </tr>
	  <script>
 var editor1;
 KindEditor.ready(function(K) {
 editor1 = K.create('textarea[name="content"]', {
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
      <td> 邮件主题 ：</td>
       <td><textarea name="content" id="content" style="width:700px;height:200px;visibility:hidden;" datatype="*" nullmsg="邮件主题不能为空"></textarea></td>
        </td>
    </tr>
    <tr>
      <td >发件人昵称：</td>
       <td>
        <input name="name" type="text" id="name" size="20"  datatype="*" nullmsg="发件人姓名不能为空"/>
      发件人姓名</td>
    </tr>
    <tr>
      <td colspan="2">
        <input type="submit" name="button" value=" 立即发送邮件 " class="bnt" />      </td>
    </tr>
	</table>
  </form>

	
<%case "pumailgo"

lmail=request("lmail")
title=request("title")
content=request("content") 
name=request("name")
	

session("title2")=title
session("content2")=content
session("name2")=name
lmail2=lmail
%>
<table width="99%" border="0" align="center" cellpadding="10" cellspacing="1" bgcolor="#f9f9f9">
    <tr>
      <td width="10" colspan="3" >
	  <iframe name="send" width="100%" height="300" marginwidth="0" marginheight="0" frameborder=no src="email_pusend.asp?lmail2=<%=lmail2%>"> 浏览器不支持嵌入式框架，或被配置为不显示嵌入式框架。</iframe></td>
    </tr>
  
	

</table>
<%End Select%>

</body>
</html>