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
<!--#include file="hqyy.asp"-->
</head>
<body>
<%select Case Request("l")%>
<%
case "somail"
%>

 <%
session("title")=""
session("content")=""
session("sendname")=""
session("sendnum")=""
session("sendtime")=""
session("hylb")=""
  %>
  <form id="form1" name="form1" method="post" action="sendhyemail.asp?l=pumailgo" class="checkform">

<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">会员邮件发送 <span id="yuyan" style="margin-left:20px"></span></div></td>
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

 <div class="dmeun" style=" background:#eeeeee; color:#999999"><a href="sendemail.asp?l=somail" target="main" style=" color:#999999">手动发送邮件</a></div>
 <div class="dmeun" style=" background:#eeeeee; color:#999999"><a href="sendmemberemail.asp?l=somail" target="main" style=" color:#999999">所有订阅用户</a></div>
 <div class="dmeun" ><a href="sendhyemail.asp?l=somail" target="main" >会员发送</a></div>
 
 </td>
 

 </tr>
</table>
  <table width="99%" border="0" align="center" cellpadding="0" cellspacing="0" class="dtable4">
  
   <tr>
     <td>会员类别：</td><td>
	 <select name="hylb" id="select" style=" border:1px #ddd solid;" datatype="*" nullmsg="请选择发送会员类别">
            <%
		  dim rsc
		  set rsc=server.CreateObject("adodb.recordset")
		  rsc.open "select * from hygroup where lang='"&lang&"'",conn,1,1
		  Response.Write("<option value="""">请选择会员类别</option>")
		  Response.Write("<option value=""1"">所有会员</option>") 
    If Rsc.Eof and Rsc.Bof Then
      
    Else
		  while not rsc.eof
			response.Write("<option value=""" & rsc("groupid") & """>" & rsc("groupname") & "</option>")
			rsc.movenext
		wend
		end if
		rsc.close
		set rsc=nothing
		  %>
                    </select></td>
   </tr>
   <tr>
      <td >邮件主题 ：</td>
	  <td>
        <input name="title" type="text" id="title" size="40"  datatype="*" nullmsg="邮件主题不能为空"/>        </td>
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
      <td >邮件内容 ：</td>
	  <td><textarea name="content" id="content" style="width:700px;height:200px;visibility:hidden;" datatype="*" nullmsg="邮件内容不能为空"></textarea></td>
    </tr>
    <tr>
      <td >每次发多少邮件：</td>
	  <td>
      <input name="sendnum" type="text" id="sendnum" value="20" size="20" maxlength="4" />
     建议不超过20封.</td>
    </tr>
    <tr>
      <td >每次发送的间隔时间：</td>
	  <td>
      <input name="sendtime" type="text" id="sendtime" value="10" size="20" maxlength="4" />
      秒  建议在10秒，以免变成垃圾邮件</td>
    </tr>
    <tr>
      <td >发件人：</td>
	  <td>
        <input name="name" type="text" id="name" size="20" datatype="*" nullmsg="发件人不能为空"/>
        发件人姓名</td>
    </tr>
    <tr>
      <td colspan="2" ><input type="submit" name="Submit52" value=" 立即发送邮件 " class="bnt" />
        此操作将发送给所有用户</td>
    </tr>
	</table>
  </form>

	
<%
case "pumailgo"

'主题
session("title")=request("title")
'内容
session("content")=Request("content")
'发送人昵称
session("sendname")=request("name")
'每次发生多少封
session("sendnum")=request("sendnum")
'发生间隔
session("sendtime")=request("sendtime")
'会员类别
session("hylb")=request("hylb")
%>
<table width="99%" border="0" align="center" cellpadding="10" cellspacing="1" bgcolor="#f9f9f9">
    <tr>
      <td width="10" colspan="3" >
	  <iframe name="send" width="100%" height="300" marginwidth="0" marginheight="0" frameborder=no src="membersendemaill.asp?pageid=1&totalcount=0"> 浏览器不支持嵌入式框架，或被配置为不显示嵌入式框架。</iframe></td>
    </tr>
  
	

</table>
<%End Select%>

</body>
</html>