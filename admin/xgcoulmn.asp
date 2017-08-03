<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<%
if Instr(session("AdminPurview"),"|10415,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</script>"
  response.end
end if
'========判断是否具有管理权限
%>
<%
If request.Form("submit") = "修改" Then
m_name=Trim(request.form("m_name")) 
m_info=Trim(request.form("m_info"))
m_id=Trim(request.form("m_id"))
m_type=Trim(request.form("m_type"))
m_sort=Trim(request.form("m_sort"))
m_sequence=Trim(request.form("m_sequence"))
if m_type=0 then
mr=request.form("m_size")
end if
if m_type=1 then
mr=request.form("m_size1")
end if
		
	if m_type=0 then
	if m_sort=1 then
	 insertSql1 = "alter table [dy] alter column "&m_name&" varchar("&mr&")"
	 end if
	if m_sort=2 then
	 insertSql1 = "alter table [news] alter column "&m_name&" varchar("&mr&")"
	 end if
	 if m_sort=3 then
	 insertSql1 = "alter table [product] alter column "&m_name&" varchar("&mr&")"
	 end if
	 if m_sort=4 then
	 insertSql1 = "alter table [al] alter column "&m_name&" varchar("&mr&")"
	 end if
	 if m_sort=5 then
	 insertSql1 = "alter table [zxsp] alter column "&m_name&" varchar("&mr&")"
	 end if
	 if m_sort=6 then
	 insertSql1 = "alter table [download] alter column "&m_name&" varchar("&mr&")"
	 end if
	 if m_sort=7 then
	 insertSql1 = "alter table [job] alter column "&m_name&" varchar("&mr&")"
	 end if
if m_sort=8 then
	 insertSql1 = "alter table [NetWork] alter column "&m_name&" varchar("&mr&")"
	 end if
	end if	
	
	
	if m_type=1 then
	
	
	if m_sort=1 then
	insertSql1 = "alter table [dy] alter column "&m_name&" Int "
	end if
	
    if m_sort=2 then
	insertSql1 = "alter table [news] alter column "&m_name&" Int "
	end if

	if m_sort=3 then
	insertSql1 = "alter table [product] alter column "&m_name&" Int "
	end if
	
	if m_sort=4 then
	insertSql1 = "alter table [al] alter column "&m_name&" Int "
	end if
	
	if m_sort=5 then
	insertSql1 = "alter table [zxsp] alter column "&m_name&" Int "
	end if
	
	if m_sort=6 then
	insertSql1 = "alter table [download] alter column "&m_name&" Int"
	end if
	
	if m_sort=7 then
	insertSql1 = "alter table [job] alter column "&m_name&" Int "
	end if
	if m_sort=8 then
	insertSql1 = "alter table [NetWork] alter column "&m_name&" Int "
	end if
	end if
	if m_type=2 Or m_type=3 then
	if m_sort=1 then
	insertSql1 = "alter table [dy] alter column "&m_name&" ntext"
	end if
	if m_sort=2 then
	insertSql1 = "alter table [news] alter column "&m_name&" ntext"
	end if
	if m_sort=3 then
	insertSql1 = "alter table [product] alter column "&m_name&" ntext"
	end if
	if m_sort=4 then
	insertSql1 = "alter table [al] alter column "&m_name&" ntext"
	end if
	if m_sort=5 then
	insertSql1 = "alter table [zxsp] alter column "&m_name&" ntext"
	end if
	if m_sort=6 then
	insertSql1 = "alter table [download] alter column "&m_name&" ntext"
	end if
	if m_sort=7 then
	insertSql1 = "alter table [job] alter column "&m_name&" ntext"
	end if
	if m_sort=8 then
	insertSql1 = "alter table [NetWork] alter column "&m_name&" ntext"
	end if
	end if	
	conn.execute insertSql1
	
    'updateSql = "ColumnName='"&m_name&"',ColumnType="&m_type&",ColumnInfo='"&m_info&"',ColumnSize='"&mr&"',Sequence="&m_sequence&",lang='"&lang&"'"
	updateSql = "ColumnName='"&m_name&"',ColumnType="&m_type&",ColumnInfo='"&m_info&"',ColumnSize='"&mr&"',Sequence="&m_sequence&""
	updateSql = "update Mycolumn set "&updateSql&" where ID="&m_id

	conn.execute updateSql

call addlog("修改字段信息")
response.Write "<script language=javascript>alert(""恭喜您修改成功"");window.location.href=""coulmn.asp""</script>"
end if

id=request.querystring("id")
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "Select * From mycolumn where id="&id&" ", conn,1,1
	
If rs.EOF Then
    response.Write "<script language=javascript>window.location.href=""coulmn.asp""</script>"
    response.End
End if
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


 <form   method="post" name="myform" >
<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">修改自定义字段  <!--<span id="yuyan" style="margin-left:20px"></span>--></div></td>
</tr>
</table>




<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td >
 <div class="dmeun"><a href="coulmn.asp" target="main">字段管理</a></div>
<div class="dmeun"><a href="coulmnadd.asp" target="main">添加字段</a></div>
 </td>
 

 </tr>
</table>
   <table width="99%"  border="0" cellpadding="0" cellspacing="0"  align="center" class="dtable4">
                  
<tr> 
<td width="11%"  > 字段名称: </td>

<td width="89%">
<input type="text" name="m_name" id="m_name"  value="<%=rs("ColumnName")%>" readonly="readonly" datatype="*" nullmsg="字段名称不能为空"/> <font color="#FF0000">*</font> 不可更改                   
</td>
</tr>
<tr>
<td  >字段说明：</td>

<td>
<input type="text" name="m_info"  value="<%=rs("ColumnInfo")%>" datatype="*" nullmsg="字段说明不能为空"/> <font color="#FF0000">*</font> 这里填写字段说明
</td>
</tr>
<tr  >
<td  >字段类型:</td>

<td>
<select name="m_type" class="input_select" onChange="if(this.value==0){$('#c_size1').show();$('#c_size2').hide()}else if(this.value==1){$('#c_size2').show();$('#c_size1').hide()}else{$('#c_size1').hide();$('#c_size2').hide()}">
<option value="0" <%if rs("ColumnType")=0 then%>selected="selected"<%end if%>>文本</option>
<option value="1" <%if rs("ColumnType")=1 then%>selected="selected"<%end if%>>数字</option>
<option value="2" <%if rs("ColumnType")=2 then%>selected="selected"<%end if%>>备注(带编辑器)</option>
<option value="2" <%if rs("ColumnType")=2 then%>selected="selected"<%end if%>>备注(普通文本框)</option>
</select>
文本或内容且对应频道已经添加了字符内容的不可更改为数字类型
</td></tr>
<tr  id="c_size1" <%if rs("ColumnType")>0 then%>style="display:none;"<%end if%>>
<td  >字段大小：</td>

<td>
<input type="text" name="m_size" style="width:80px;" value="<%=rs("ColumnSize")%>"/>这里填写小于255的数字
</td>
</tr>


				  
				 <tr  > 
                    <td  > 字段排序:</td>

<td>
                     <input type="text" name="m_sequence"  style="width:80px;" value="<%=rs("Sequence")%>" /></td>
                  </tr>
				 <input type="hidden" name="m_id" value="<%=id%>">
    <input type="hidden" name="m_sort" value="<%=rs("TypeID")%>">
                  <tr  > 
                    <td colspan="2"> 
 <input class="bnt" name="submit" type="submit" value="修改">
&nbsp;                                      </td>
                  </tr>
   </table>
            
</form>
   
</body>
</html>
