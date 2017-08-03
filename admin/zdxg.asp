<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="Inc/Function.asp"-->
<!--#include file="admin.asp"-->
<% 
diyform=request.QueryString("form")
id=request.QueryString("id")
id1=request.QueryString("id1")
If request.Form("submit") = "修改" Then
Alias=Trim(request.form("Alias")) 
IsNotNull=Trim(request.form("IsNotNull"))
ListVal=Trim(request.form("ListVal"))
FieldNull=Trim(request.form("FieldNull"))
xs=Trim(request.form("xs"))
Px=Trim(request.form("Px"))
YanZheng=Trim(request.form("YanZheng"))
Set rs = server.CreateObject("adodb.recordset")
rs.open "Select * From FreeField where id="&id&" ",conn,1,3
rs("Alias")=Alias
rs("xs")=xs
rs("IsNotNull")=IsNotNull
rs("ListVal")=ListVal
rs("FieldNull")=FieldNull
rs("YanZheng")=YanZheng
rs("lang")=lang
rs("px")=px
rs.update
rs.close
set rs=nothing


call addlog("修改自定义字段")
response.Write "<script language=javascript>alert(""恭喜您修改成功"");window.location.href=""zd.asp?form="&diyform&"&id="&id1&"""</script>"
end if
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "Select * From FreeField where id="&id&" and lang='"&lang&"'", conn,1,1


	
If rs.EOF Then
    response.Write "<script language=javascript>window.location.href=""bd.asp""</script>"
    response.End
End if
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
<script language="JavaScript" src="js/jquery-1.7.2.min.js"></script>
<style type="text/css">
.table1{ font-size: 14px; background: #F9f9f9;border:none;}
.table1 td{padding:10px; font-family:"微软雅黑";color:#ccc; font-weight:bold;}
.table1 td a{color:#ffffff}
</style>
<!--#include file="hqyy.asp"-->
</style>
</head>

<body>
<form method="post" name="myform"  class="checkform">

<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">修改自定义字段 <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>
<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="chajian.asp" target="main">返回菜单</a></div><div class="dmeun"><a href="bd.asp" target="main">自定义表单管理</a></div><div class="dmeun"><a href="zd.asp?form=<%=diyform%>&id=<%=id1%>" target="main">字段管理</a></div></td>
 

 </tr>
</table>

   <table width="99%"  border="0" cellpadding="0" cellspacing="0"  align="center" class="dtable4">
        <col width="200px">
          </col>
        
        <tbody>
          <tr>
            <td>字段别名：</td>
			
			
            <td ><input type="text" name="Alias" id="Alias"  autocomplete="off" datatype="*" nullmsg="字段别名必须填写"  value="<%=rs("Alias")%>"/>
              <span class="Validform_checktip">*,如：姓名</span></td>
          </tr>
          <!--
          <tr>
            <td>字段名称：</td>
            <td ><input type="text" name="FieldName" id="FieldName"  autocomplete="off" value="<%=rs("FieldName")%>" datatype="*" nullmsg="字段名称必须填写" readonly/>
              <span class="Validform_checktip">* 字段名称不能修改，如果需要，请重新添加</span></td>
          </tr>
          -->

       
          <tr>
            <td>是否必填：</td>
            <td ><span id="cblItem">
              <input  type="radio" name="IsNotNull" value="1"   <%if rs("IsNotNull")="1" then response.write "checked" end if%>/>
              <label >是</label>
              <input  type="radio" name="IsNotNull" value="0"  <%if rs("IsNotNull")="0" then response.write "checked" end if%>/>
              <label >否</label>
              </span><span class="Validform_checktip"></span></td>
          </tr>
          <!--
          <tr>
            <td>字段类型：</td>
            <td ><select size="1" name="FieldType" readonly>
        <option value="1" <%if rs("FieldType")="1" then response.write ("selected")%>>文本框</option>
        <option value="2" <%if rs("FieldType")="2" then response.write ("selected")%>>单选</option>
        <option value="3" <%if rs("FieldType")="3" then response.write ("selected")%>>多选</option>
        <option value="4" <%if rs("FieldType")="4" then response.write ("selected")%>>列表框</option>
        <option value="5" <%if rs("FieldType")="5" then response.write ("selected")%>>备注型</option>
		<option value="6" <%if rs("FieldType")="6" then response.write ("selected")%>>时间</option>
                                                </select></td>
          </tr>
          -->

<tr>
				<td>
			高级验证：</td>
			
			<td>
			<select size="1" name="yanzheng">    
              <option value="0" <%if rs("yanzheng")="0" then response.write ("selected")%>>无限制</option>
              <option value="1" <%if rs("yanzheng")="1" then response.write ("selected")%>>email</option>
              <option value="2" <%if rs("yanzheng")="2" then response.write ("selected")%>>身份证</option>
			<option value="3" <%if rs("yanzheng")="3" then response.write ("selected")%>>必须为数字</option>
			<option value="4" <%if rs("yanzheng")="4" then response.write ("selected")%>>手机号码</option>
			<option value="5" <%if rs("yanzheng")="5" then response.write ("selected")%>>邮政编码</option>
            </select>
					  
					  
					高级验证 必须是 必选条件下有效  </td>
				  </tr>



          



 <tr>
				    <td  >单选和多选和列表框的值：</td>
					<td>
				      <label>
				      <input name="ListVal" type="text" id="ListVal" size="80" value="<%=rs("ListVal")%>"/>
					  <br /><br />
			        当类型选择为单选或者是多选时请填写，其他请留空,多个值之间请用|隔开。 当为列表框时：请选择性别|,男|男,女|女</label></td>
    </tr>






          
          <tr>
            <td>提示信息显示：</td>
            <td><input type="text" name="FieldNull" id="FieldNull"   value="<%=rs("FieldNull")%>">
              </td>
          </tr>
 

   <tr>
            <td>后台信息列表页是否显示：</td>
            <td >
              <input  type="radio" name="xs" value="1"   <%if rs("xs")="1" then response.write "checked" end if%>/>
              是
              <input  type="radio" name="xs" value="0"  <%if rs("xs")="0" then response.write "checked" end if%>/>
              否
             </td>
          </tr>


          
          <tr>
            <td> 显示顺序： </td>
            <td ><input type="text"  name="Px" id="Px" autocomplete="off" value="<%=rs("px")%>" datatype="*" nullmsg="排序必须填写"  />
              <span class="Validform_checktip">*</span></td>
          </tr>
          <tr>
            <td colspan=2>
           <input name="submit" type="submit" value="修改" class="bnt"/>
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

