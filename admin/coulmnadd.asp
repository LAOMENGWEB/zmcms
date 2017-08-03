<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<%
if Instr(session("AdminPurview"),"|10415,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</script>"
response.end
end if
'========判断是否具有管理权限
If request.form("submit") = "确定" Then
m_name=Trim(request.form("m_name")) 
m_info=Trim(request.form("m_info"))
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
if sqlno=1 then
if m_sort=1 then
insertSql1 = "alter table [dy] add column "&m_name&" varchar("&mr&")"
end if
if m_sort=2 then
insertSql1 = "alter table [news] add column "&m_name&" varchar("&mr&")"
end if
if m_sort=3 then
insertSql1 = "alter table [product] add column "&m_name&" varchar("&mr&")"
end if
if m_sort=4 then
insertSql1 = "alter table [al] add column "&m_name&" varchar("&mr&")"
end if
if m_sort=5 then
insertSql1 = "alter table [zxsp] add column "&m_name&" varchar("&mr&")"
end if
if m_sort=6 then
insertSql1 = "alter table [download] add column "&m_name&" varchar("&mr&")"
end if
if m_sort=7 then
insertSql1 = "alter table [job] add column "&m_name&" varchar("&mr&")"
end if
if m_sort=8 then
insertSql1 = "alter table [NetWork] add column "&m_name&" varchar("&mr&")"
end if
end if
if sqlno=2 then
if m_sort=1 then
insertSql1 = "alter table [dy] add  "&m_name&" varchar("&mr&")"
end if
if m_sort=2 then
insertSql1 = "alter table [news] add  "&m_name&" varchar("&mr&")"
end if
if m_sort=3 then
insertSql1 = "alter table [product] add  "&m_name&" varchar("&mr&")"
end if
if m_sort=4 then
insertSql1 = "alter table [al] add  "&m_name&" varchar("&mr&")"
end if
if m_sort=5 then
insertSql1 = "alter table [zxsp] add  "&m_name&" varchar("&mr&")"
end if
if m_sort=6 then
insertSql1 = "alter table [download] add  "&m_name&" varchar("&mr&")"
end if
if m_sort=7 then
insertSql1 = "alter table [job] add  "&m_name&" varchar("&mr&")"
end if
if m_sort=8 then
insertSql1 = "alter table [NetWork] add  "&m_name&" varchar("&mr&")"
end if
end if
end if	
if m_type=1 then
if sqlno=1 then
if m_sort=1 then
insertSql1 = "alter table [dy] add column "&m_name&" Int "
end if
if m_sort=2 then
insertSql1 = "alter table [news] add column "&m_name&" Int "
end if
if m_sort=3 then
insertSql1 = "alter table [product] add column "&m_name&" Int "
end if
if m_sort=4 then
insertSql1 = "alter table [al] add column "&m_name&" Int "
end if
if m_sort=5 then
insertSql1 = "alter table [zxsp] add column "&m_name&" Int"
end if
if m_sort=6 then
insertSql1 = "alter table [download] add column "&m_name&" Int"
end if
if m_sort=7 then
insertSql1 = "alter table [job] add column "&m_name&" Int "
end if
if m_sort=8 then
insertSql1 = "alter table [NetWork] add column "&m_name&" Int "
end if
end if
if sqlno=2 then
if m_sort=1 then
insertSql1 = "alter table [dy] add  "&m_name&" Int "
end if
if m_sort=2 then
insertSql1 = "alter table [news] add  "&m_name&" Int "
end if
if m_sort=3 then
insertSql1 = "alter table [product] add  "&m_name&" Int "
end if
if m_sort=4 then
insertSql1 = "alter table [al] add  "&m_name&" Int "
end if
if m_sort=5 then
insertSql1 = "alter table [zxsp] add  "&m_name&" Int"
end if
if m_sort=6 then
insertSql1 = "alter table [download] add  "&m_name&" Int"
end if
if m_sort=7 then
insertSql1 = "alter table [job] add  "&m_name&" Int"
end if
if m_sort=8 then
insertSql1 = "alter table [NetWork] add  "&m_name&" Int"
end if
end if
end if
if m_type=2 Or m_type=3 then
if sqlno=1 then
if m_sort=1 then
insertSql1 = "alter table [dy] add column "&m_name&" ntext"
end if
if m_sort=2 then
insertSql1 = "alter table [news] add column "&m_name&" ntext"
end if
if m_sort=3 then
insertSql1 = "alter table [product] add column "&m_name&" ntext"
end if
if m_sort=4 then
insertSql1 = "alter table [al] add column "&m_name&" ntext"
end if
if m_sort=5 then
insertSql1 = "alter table [zxsp] add column "&m_name&" ntext"
end if
if m_sort=6 then
insertSql1 = "alter table [download] add column "&m_name&" ntext"
end if
if m_sort=7 then
insertSql1 = "alter table [job] add column "&m_name&" ntext"
end if
if m_sort=8 then
insertSql1 = "alter table [NetWork] add column "&m_name&" ntext"
end if
end if
if sqlno=2 then
if m_sort=1 then
insertSql1 = "alter table [dy] add  "&m_name&" ntext"
end if
if m_sort=2 then
insertSql1 = "alter table [news] add  "&m_name&" ntext"
end if
if m_sort=3 then
insertSql1 = "alter table [product] add  "&m_name&" ntext"
end if
if m_sort=4 then
insertSql1 = "alter table [al] add  "&m_name&" ntext"
end if
if m_sort=5 then
insertSql1 = "alter table [zxsp] add  "&m_name&" ntext"
end if
if m_sort=6 then
insertSql1 = "alter table [download] add  "&m_name&" ntext"
end if
if m_sort=7 then
insertSql1 = "alter table [job] add  "&m_name&" ntext"
end if
if m_sort=8 then
insertSql1 = "alter table [NetWork] add  "&m_name&" ntext"
end if
end if
end if	
conn.execute insertSql1
'insertSql = "insert into [Mycolumn](TypeID,ColumnName,ColumnType,ColumnInfo,ColumnSize,Sequence,lang) values ("&m_sort&",'"&m_name&"',"&m_type&",'"&m_info&"','"&mr&"',"&m_sequence&",'"&lang&"')"
insertSql = "insert into [Mycolumn](TypeID,ColumnName,ColumnType,ColumnInfo,ColumnSize,Sequence) values ("&m_sort&",'"&m_name&"',"&m_type&",'"&m_info&"','"&mr&"',"&m_sequence&")"

conn.execute insertSql
call addlog("添加字段信息")
response.Write "<script language=javascript>alert(""恭喜您添加成功"");window.location.href=""coulmn.asp""</script>"
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

 <form   method="post" name="myform" class="checkform">
 
 
 
<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">添加字段  <!--<span id="yuyan" style="margin-left:20px"></span>--></div></td>
</tr>
</table>




<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td >
 <div class="dmeun"><a href="coulmn.asp" target="main">字段管理</a></div>

 </td>
 

 </tr>
</table>
 
 
 
 
 
   <table width="99%"  border="0" cellpadding="0" cellspacing="0"  align="center" class="dtable4">
                  
<tr> 
<td width="9%"  > 字段名称:</td>
<td width="91%">
<input type="text" name="m_name" id="m_name"  value="Z_" datatype="*" nullmsg="字段名称不能为空"/> <font color="#FF0000">*</font> 添加后不可更改 尽量用英文   </font>               
</td>
</tr>
<tr>
<td  >字段说明：</td>
<td>
<input type="text" name="m_info"  datatype="*" nullmsg="字段说明不能为空"/> <font color="#FF0000">*</font> 这里填写字段说明
</td>
</tr>
<tr  >
<td  >字段类型:</td>
<td>
<select name="m_type"  onChange="if(this.value==0){$('#c_size1').show();$('#c_size2').hide()}else if(this.value==1){$('#c_size2').show();$('#c_size1').hide()}else{$('#c_size1').hide();$('#c_size2').hide()}">
<option value="0">文本</option>
<option value="1">数字</option>
<option value="2">备注（带编辑器）</option>
<option value="3">备注（普通备注框）</option>
</select>
</td>
</tr>
<tr  id="c_size1">
<td  >字段大小：</td>
<td>
<input type="text" name="m_size"  value="200" style="width:80px;" /> 这里填写小于255的数字
</td>
</tr>

<tr>
<td  >所属频道：</td>
<td>
<select name="m_sort" class="input_select">
            <option value="1"><%=danye%></option>
            <option value="2"><%=xinwen%></option>
            <option value="3"><%=chanpin%></option>
            <option value="4"><%=tupian%></option>
            <option value="5"><%=shipin%></option>
            <option value="6"><%=xiazai%></option>
            <option value="7"><%=zhaopin%></option>
			<option value="8"><%=yingxiao%></option>
        </select> <font color="#FF0000">*</font> 选择后不能更改</td>
			      </tr>			  
				  
				 <tr  > 
                    <td  > 字段排序:</td>
<td>
                     <input type="text" name="m_sequence"  style="width:80px;" value="1" /></td>
                  </tr>
				 	
                  <tr  > 
                    <td  colspan="2"> 
 <input class="bnt" name="submit" type="submit" value="确定">
&nbsp;                                      </td>
                  </tr>
   </table>
            
</form>
   
</body>
</html>
