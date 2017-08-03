<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="Inc/Function.asp"-->
<!--#include file="admin.asp"-->
<%
id=request.QueryString("id")
If request.Form("submit") = "修改" Then
IsStatus=Trim(request.form("IsStatus")) 
FormName=Trim(request.form("FormName"))
SuccessMsg=Trim(request.form("SuccessMsg"))
IsCode=Trim(request.form("IsCode"))
zxs=Trim(request.form("zxs"))
txs=Trim(request.form("txs"))
fsh=Trim(request.form("fsh"))
rs("txs")=txs
Set rs = server.CreateObject("adodb.recordset")
rs.open "Select * From FreeForm where id="&id&" ",conn,1,3
rs("IsStatus")=IsStatus
rs("FormName")=FormName
rs("SuccessMsg")=SuccessMsg
rs("IsCode")=IsCode
rs("zxs")=zxs
rs("txs")=txs
rs("fsh")=fsh
rs("lang")=lang
rs.update
rs.close
set rs=nothing
call addlog("修改自定义表单")
response.Write "<script language=javascript>alert(""恭喜您修改成功"");window.location.href=""bd.asp""</script>"
end if

Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "Select * From FreeForm where id="&id&" and lang='"&lang&"'", conn,1,1
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
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">修改自定义表单  <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>

<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="chajian.asp" target="main">返回菜单</a></div><div class="dmeun"><a href="bd.asp" target="main">自定义表单管理</a></div></td>
 

 </tr>
</table>

   <table width="99%"  border="0" cellpadding="0" cellspacing="0"  align="center" class="dtable4">
                  
	 <tr  > 
                    <td width="12%"  > 表单状态:  </td>
       <td width="88%" colspan="2">
	   
	     <input type="radio" name="IsStatus"  value="1" <%if rs("IsStatus")="1" then response.write "checked"%>/>
              正常
         <input type="radio" name="IsStatus"  value="0" <%if rs("IsStatus")="0" then response.write "checked"%>/>
         停用

	   
	   
	   </td>
                    
	 </tr> 
					 
				  <tr  >
				    <td  >表单名称:</td>
			        <td  > <input type="text"  name="FormName" id="FormName" datatype="*" nullmsg="表单名称必须填写"  value="<%=rs("FormName")%>" />
</td>
			        <td  ><div class="Validform_checktip">表单名称必须填写</div></td>
				  </tr>
				  <tr  >
				    <td  > 数据表： </td>
			        <td ><input type="text" class="txtInput w150" name="FormTableName" id="FormTableName"  value="<%=rs("FormTableName")%>" readonly/>
</td>
					<td>数据表不能修改</td>
				  </tr>
				  <tr  > 
                    <td  > 表单提交成功提示： </td>
                    <td  colspan="2"><input type="text"  name="SuccessMsg" id="SuccessMsg"  datatype="*" nullmsg="表单提交成功提示必须填写"  value="<%=rs("SuccessMsg")%>" />
</td>
				  </tr>
				  

				       <tr>
            <td>字段别名是否显示：</td>
			
			
            <td colspan="2">

            <input  type="radio" name="zxs" value="1" <%if rs("zxs")="1" then response.write "checked" end if%>/>
            是
              <input  type="radio" name="zxs" value="0"  <%if rs("zxs")="0" then response.write "checked" end if%>/>
            否

	            
            </td>
          </tr>
				  
				   <tr>
            <td>提示信息显示位置：</td>
			
			
            <td colspan="2">

            <input  type="radio" name="txs" value="1" <%if rs("txs")="1" then response.write "checked" end if%>/>
             输入框后面
              <input  type="radio" name="txs" value="0"  <%if rs("txs")="0" then response.write "checked" end if%>/>
              输入框默认值

	            
            </td>
          </tr>  
				  
				 			  	<tr> 
      <td >信息是否需要审核： </td>
      <td colspan="2"><input type="radio" class="none"  name="fsh" value="1" <%if rs("fsh")="1" then response.write "checked"%>>
  是 <input type="radio" class="none"  name="fsh" value="2" <%if rs("fsh")="2" then response.write "checked"%>>
      否</td>
	</tr>
	
                    <td  > 是否开启验证码：</td>
                    <td  colspan="2">
					  <input type="radio" name="IsCode"  value="1" checked="checked" datatype="*"  <%if rs("IsCode")="1" then response.write "checked"%>/>
              正常
         <input type="radio" name="IsCode" id="IsCode" value="0" <%if rs("IsCode")="0" then response.write "checked"%>/>
         停用
					
					
					
					
					</td>
     </tr>
                  <tr  > 
                    <td colspan="3" style="border-bottom:none"> 
                        <input class="bnt" name="submit" type="submit" value="修改">
&nbsp;                                      </td>
                  </tr>
   </table>
            
</form>
   

</body>
</html>
