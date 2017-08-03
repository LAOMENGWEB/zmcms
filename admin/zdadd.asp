<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="Inc/Function.asp"-->
<!--#include file="admin.asp"-->

<% 
diyform=request.QueryString("form")
ChannelId=request.QueryString("id")
If request.Form("submit") = "确定" Then
Alias=Trim(request.form("Alias")) 
FieldName=Trim(request.form("FieldName"))
IsNotNull=Trim(request.form("IsNotNull"))
xs=Trim(request.form("xs"))
zxs=Trim(request.form("zxs"))
txs=Trim(request.form("txs"))
FieldType=Trim(request.form("FieldType"))
ListVal=Trim(request.form("ListVal"))
FieldNull=Trim(request.form("FieldNull"))
Px=Trim(request.form("Px"))
YanZheng=Trim(request.form("YanZheng"))

if FieldType=1 or FieldType=2 or FieldType=3 or FieldType=4  then
if IsFieldsText(""&diyform&"",""&FieldName&"")=true Then
 response.Write "<script language=javascript>alert(""非常抱歉！此数据表字段已存在于"&diyform&"表中！请更换字段名"");window.location.href=""zd.asp?form="&diyform&"&id="&ChannelId&"""</script>"
 response.End()	
else
AddColumn ""&diyform&"",""&FieldName&"","nvarchar(255) Null"
end if
end if

if FieldType=5  then
	if IsFieldsText(""&diyform&"",""&FieldName&"")=true Then
 response.Write "<script language=javascript>alert(""非常抱歉！此数据表字段已存在于"&diyform&"表中！请更换字段名"");window.location.href=""zd.asp?form="&diyform&"&id="&ChannelId&"""</script>"
 response.End()	
 else
AddColumn ""&diyform&"",""&FieldName&"","ntext"
end if
end if

if FieldType=6  then
	if IsFieldsText(""&diyform&"",""&FieldName&"")=true Then
 response.Write "<script language=javascript>alert(""非常抱歉！此数据表字段已存在于"&diyform&"表中！请更换字段名"");window.location.href=""zd.asp?form="&diyform&"&id="&ChannelId&"""</script>"
 response.End()	
 else
AddColumn ""&diyform&"",""&FieldName&"","DATETIME  Null"
end if
end if
Set rs = server.CreateObject("adodb.recordset")
rs.open "Select * From FreeField ",conn,1,3
rs.addnew
rs("Alias")=Alias
rs("FieldName")=FieldName
rs("IsNotNull")=IsNotNull
rs("FieldType")=FieldType
rs("ListVal")=ListVal
rs("xs")=xs
rs("lang")=lang
rs("FieldNull")=FieldNull
rs("YanZheng")=YanZheng
rs("ChannelId")=ChannelId
rs("px")=px
rs("addtime")=now()
rs.update
rs.close
set rs=nothing


call addlog("添加自定义字段")
response.Write "<script language=javascript>alert(""恭喜您添加成功"");window.location.href=""zd.asp?form="&diyform&"&id="&ChannelId&"""</script>"
end if
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>zmcms_管理页面</title>
<link href="css/style.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="js/jquery-1.7.2.min.js"></script>
 <script type="text/javascript" src="../js/Validform_v5.3.2_min.js"></script>
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
<form method="post" name="myform"  class="checkform">



<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">添加自定义字段  <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>
<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="chajian.asp" target="main">返回菜单</a></div><div class="dmeun"><a href="bd.asp" target="main">自定义表单管理</a></div><div class="dmeun"><a href="zd.asp?form=<%=diyform%>&id=<%=ChannelId%>" target="main">字段管理</a></div></td>
 

 </tr>
</table>

   <table width="99%"  border="0" cellpadding="0" cellspacing="0"  align="center" class="dtable4">
        <col width="200px">
          </col>
        
        <tbody>
          <tr>
            <td>字段别名：</td>
			
			
            <td ><input type="text" name="Alias" id="Alias"  autocomplete="off" datatype="*" nullmsg="字段别名必须填写"  />
              <span class="Validform_checktip">*,如：姓名</span></td>
          </tr>

  




          
          <tr>
            <td>字段名称：</td>
            <td ><input type="text" name="FieldName" id="FieldName"  autocomplete="off" value="<%=FieldPrefix%>" datatype="*" nullmsg="字段名称必须填写"  />
              <span class="Validform_checktip">*</span></td>
          </tr>
          <tr>
            <td>是否必填：</td>
            <td >
              <input  type="radio" name="IsNotNull" value="1"  />
              是
              <input  type="radio" name="IsNotNull" value="0"  checked/>
              否
              </td>
          </tr>
          <tr>
            <td>字段类型：</td>
            <td ><select size="1" name="FieldType">
        <option value="1" selected="selected">文本框</option>
        <option value="2">单选</option>
        <option value="3">多选</option>
        <option value="4">列表框</option>
        <option value="5">备注型</option>
		<option value="6">时间</option>
                                                </select></td>
          </tr>

<tr>
				<td>
			高级验证：</td>
			
			<td>
			<select size="1" name="yanzheng">    
              <option value="0">无限制</option>
              <option value="1">email</option>
              <option value="2">身份证</option>
			<option value="3">必须为数字</option>
			<option value="4">手机号码</option>
			<option value="5">邮政编码</option>
            </select>
					  
					  
					高级验证 必须是 必选条件下有效  </td>
				  </tr>



          



 <tr>
				    <td  >单选和多选和列表框的值：</td>
					<td>
				      <label>
				      <input name="ListVal" type="text" id="ListVal" size="80" />
					  <br /><br />
			        当类型选择为单选或者是多选时请填写，其他请留空,多个值之间请用|隔开。 当为列表框时：请选择性别|,男|男,女|女</label></td>
    </tr>






          
          <tr>
            <td>提示信息显示：</td>
            <td><input type="text" name="FieldNull" id="FieldNull" />
              </td>
          </tr>
          


     <tr>
            <td>后台信息列表是否显示：</td>
            <td >
              <input  type="radio" name="xs" value="1" checked="checked"/>
             是
              <input  type="radio" name="xs" value="0"  />
              否
             </td>
          </tr>



          
          <tr>
            <td> 显示顺序： </td>
            <td ><input type="text"  name="Px" id="Px" autocomplete="off" value="1" datatype="*" nullmsg="排序必须填写"  />
              <span class="Validform_checktip">*</span></td>
          </tr>
          <tr>
            <td colspan=2>
           <input name="submit" type="submit" value="确定" class="bnt"/>
       </td>
          </tr>
        </tbody>
      </table>
             
</form>
 

</body>
</html>
<%
rs.close
set rs=nothing
%>

