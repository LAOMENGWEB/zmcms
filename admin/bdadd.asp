<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="Inc/Function.asp"-->
<%
If request.Form("submit") = "确定" Then
IsStatus=Trim(request.form("IsStatus")) 
FormName=Trim(request.form("FormName"))
FormTableName=Trim(request.form("FormTableName"))
SuccessMsg=Trim(request.form("SuccessMsg"))
IsCode=Trim(request.form("IsCode"))
zxs=Trim(request.form("zxs"))
txs=Trim(request.form("txs"))
fsh=Trim(request.form("fsh"))
CreateTable(""&FormTableName&"")

Set rs = server.CreateObject("adodb.recordset")
rs.open "Select * From FreeForm ",conn,1,3
rs.addnew
rs("IsStatus")=IsStatus
rs("FormName")=FormName
rs("FormTableName")=FormTableName
rs("SuccessMsg")=SuccessMsg
rs("IsCode")=IsCode
rs("zxs")=zxs
rs("txs")=txs
rs("fsh")=fsh
rs("adddate")=now()
rs("lang")=lang
rs.update
rs.close
set rs=nothing
AddColumn ""&FormTableName&"", "Ip", "nvarchar(50) Null"
AddColumn ""&FormTableName&"", "lang", "nvarchar(50) Null"
AddColumn ""&FormTableName&"", "replay", "ntext"
AddColumn ""&FormTableName&"", "replaytime", "nvarchar(50) Null"
AddColumn ""&FormTableName&"", "sh", "int null"
AddColumn ""&FormTableName&"", "Adddate", "DATETIME Null " 
call addlog("添加自定义表单")
response.Write "<script language=javascript>alert(""恭喜您添加成功"");window.location.href=""bd.asp""</script>"
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
 <form  method="post" name="myform"  class="checkform">


<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">添加自定义表单  <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>
<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="chajian.asp" target="main">返回菜单</a></div><div class="dmeun"><a href="dy.asp" target="main">自定义表单管理</a></div></td>
 

 </tr>
</table>

   <table width="99%"  border="0" cellpadding="0" cellspacing="0"  align="center" class="dtable4">
                  
	 <tr  > 
                    <td width="12%"  > 表单状态:  </td>
       <td width="88%" colspan="2">
	   
	     <input type="radio" name="IsStatus"  value="1" checked="checked" datatype="*" nullmsg="请选择表单状态！" />
              正常
         <input type="radio" name="IsStatus" id="IsStatus0" value="0" />
         停用

	   
	   
	   </td>
                    
	 </tr> 
					 
				  <tr  >
				    <td  >表单名称:</td>
			        <td  > <input type="text"  name="FormName" id="FormName" datatype="*2-20" nullmsg="表单名称必须填写" errormsg="至少2个字符,最多20个字符!" />
</td>
			        <td  ><div class="Validform_checktip">表单名称必须填写</div></td>
				  </tr>
				  <tr  >
				    <td  > 数据表： </td>
			        <td ><input type="text" class="txtInput w150" name="FormTableName" id="FormTableName"  />
</td>
					<td>数据表不能为空</td>
				  </tr>
				  <tr  > 
                    <td  > 表单提交成功提示： </td>
                    <td  colspan="2"><input type="text"  name="SuccessMsg" id="SuccessMsg"  datatype="*2-100" nullmsg="表单提交成功提示必须填写" errormsg="至少2个字符,最多100个字符!" />
</td>
				  </tr>
				  
				  
				  
				    <tr>
            <td>字段别名前台是否显示：</td>
			
			
            <td colspan="2">

            <input  type="radio" name="zxs" value="1" checked="checked"/>
             是
              <input  type="radio" name="zxs" value="0"  />
              否

	            
            </td>
          </tr>
              <tr>
            <td>字段提示信息显示位置：</td>
			
			
            <td colspan="2">

            <input  type="radio" name="txs" value="1" checked="checked"/>
             输入框后面
              <input  type="radio" name="txs" value="0"  />
              输入框默认值

	            
            </td>
          </tr>
				  
	
				 
				  
				  	<tr> 
      <td >信息是否需要审核： </td>
      <td colspan="2"><input type="radio" class="none"  name="fsh" value="1" >
  是 <input type="radio" class="none"  name="fsh" value="2" checked="checked">
      否</td>
	</tr>
				  
				 
	
                    <td  > 是否开启验证码：</td>
                    <td  colspan="2">
					  <input type="radio" name="IsCode"  value="1" checked="checked" datatype="*"  />
              正常
         <input type="radio" name="IsCode" id="IsCode" value="0" />
         停用
					
					
					
					
					</td>
     </tr>
                  <tr  > 
                    <td colspan="3" style="border-bottom:none"> 
                        <input class="bnt" name="submit" type="submit" value="确定">
&nbsp;                                      </td>
                  </tr>
   </table>
            
</form>
   

</body>
</html>
