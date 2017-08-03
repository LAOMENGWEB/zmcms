<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<% 
id=Request.QueryString("id")
id1=Request.QueryString("id1")
form=Request.QueryString("form")
if id="" or not isnumeric(id) then
Response.Write "<script>alert('参数错误！');history.go(-1);</script>" 
Response.End()
end if
If request.Form("submit") = "回复" Then 
replay=Trim(request.form("replay"))
sh=Trim(request.form("sh"))

Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from "&form&" where id="&id
rs.open sql,conn,1,3
rs("replay") = replay
rs("replaytime") = now()
if sh="1" then
		rs("sh")=1
	else
		rs("sh")=0
	end if
rs.update
call addlog("回复表单反馈信息")
response.Write "<script language=javascript>alert(""恭喜您回复成功"");window.location.href=""fankui.asp?form="&form&"&id="&id1&"""</script>"
end if
rs.close
set rs=nothing
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>zmcms_管理页面</title>
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
		
	   <script>
 var editor;
 KindEditor.ready(function(K) {
 editor = K.create('textarea[name="replay"]', {
 filterMode : false, 
 uploadJson : '../zmbjq/asp/upload_json.asp',
 fileManagerJson : '../zmbjq/asp/file_manager_json.asp',
 allowFileManager : true,
 width: '700px',
 hight: '300px',
 afterBlur: function(){this.sync();}
 
 });
 	K('input[name=appendHtml]').click(function(e) {
					editor.insertHtml('[zmcmsfy]');
				});
			
			prettyPrint();
 });
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
rs.Open "Select * From "&form&" where id="&id&" and lang='"&lang&"'", conn,1,1
If rs.EOF Then
    response.Write "<script language=javascript>window.location.href=""bd.asp""</script>"
    response.End
End if
%>	
<form method="post" name="myform" class="checkform">


<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">回复信息 <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>
<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="fankui.asp?form=<%=form%>&id=<%=id1%>" target="main">反馈信息管理</a></div>


<div class="dmeun"><a href="bd.asp" target="main">自定义表单管理</a></div>
 </td>
 </tr>
</table>

<table width="99%"  border="0" cellpadding="0" cellspacing="0" align="center" class="dtable4">
<%
set rs1=server.createobject("adodb.recordset")
sql="select  * from FreeField where ChannelId="&id1&" and lang='"&lang&"' order by px asc"
rs1.open sql,conn,1,3
do while not rs1.eof
%>
<tr >
<td width="11%" ><%=rs1("Alias")%>：</td>
<td ><%=rs(""&rs1("FieldName")&"")%></td>
</tr>
<%
rs1.movenext
loop%>

<%
rs1.close
set rs1=nothing
%>
<tr >
<td width="11%" >反馈IP：</td>
<td><%=rs("ip")%></td>
</tr>
<tr >
<td width="11%" >反馈时间：</td>
<td><%=rs("adddate")%></td>
</tr>

<%if rs("replaytime")<>"" then%>
<tr >
<td width="11%" >回复时间：</td>
<td><%=rs("replaytime")%></td>
</tr>
<%end if%>
<tr >
<td width="11%" >回复内容：</td>
<td><textarea name="replay" style="width:80%;height:200px;visibility:hidden;" datatype="*" nullmsg="回复内容不能为空"><%=rs("replay")%></textarea></td>
</tr>
<tr >
 <td  >是否审核通过并在前台显示：</td>
 <td><input name="sh" class="none" type="checkbox" id="sh" value="1" <% if rs("sh")=1 then response.Write("checked") end if%>></td>
</tr>
 <tr >
<td   colspan="2">
<div align="left"><input name="submit" type="submit" value="回复" class="bnt"></div>
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
