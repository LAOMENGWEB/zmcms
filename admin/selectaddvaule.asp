<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!-- #include file="Inc/Head.asp" -->
<!--#include file="Inc/Function.asp"-->
<%
if Instr(session("AdminPurview"),"|36,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</script>"
  response.end
end if
'========判断是否具有管理权限
%>
<%
id=request.QueryString("id")
mode=request.QueryString("mode")
Set rs3 = server.CreateObject("adodb.recordset")
rs3.open "Select * From SelectName Where s_id=" & id & "",conn,1,3
typeid=rs3("typeid")
set rs3=nothing
rs3.close


If request.Form("submit") = "提交" Then
s_value=Trim(request.form("s_value"))
s_order=trim(request.form("s_order"))
Set rs = server.CreateObject("adodb.recordset")
rs.open "Select * From SelectValue ",conn,1,3

rs.addnew
rs("s_value")=s_value
rs("s_order")=s_order
rs("s_mode")=mode
rs("s_id")=id
rs("typeid")=typeid
rs.update
rs.close
set rs=nothing 
call addlog("添加附属属性类别")

response.Write "<script language=javascript>alert(""恭喜您添加成功"");window.location.href=""selectglvaule.asp?mode="&mode&"&id="&id&"""</script>"


end if

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>附属属性类别增加</title>
<style type="text/css">
table{ border:1px solid #D4D4D4; line-height:26px;font-family: Arial, Helvetica, sans-serif;font-size: 12px; background-color:#faf9d1}
table td{BORDER-bottom: #D4D4D4 1px dashed;}
.ys td{ padding:10px}
 .imgDiv { float:left; padding-top:5px; padding-left:5px; }
 .imgDiv img{ border:0px; width:120px; height:90px}
</style>
<SCRIPT language=javaScript>
function Checkly()
{
	if (document.myform.s_name.value==""){
		alert ("提示：\n\n类别名称不能为空");
		document.myform.s_name.focus();
		return false;
	}
		  

      

 }
</SCRIPT>
</head>

<body>
<table width="99%" border="0" align="center" cellpadding="0" cellspacing="0" style="border-bottom:none">
  <tr >
     <td  align="center"  style="border:none" class="title"><a href="selectvauleadd.asp" target="main" class="white">添加属性值</a></td>
    <td align="center"  style="border:none" class="title2"><a href="selectvaule.asp" target="main" class="white">附属属性类别管理</a></td>
 
   <td  align="center"  style="border:none" class="title"><a href="select.asp?mode=0" target="main" class="white">筛选属性管理</a></td>
	<td  align="center"  style="border:none" class="title2"><a href="select.asp?mode=2" target="main" class="white">产品属性管理</a></td>
    	<td  align="center"  style="border:none" class="title"><a href="select.asp?mode=1" target="main" class="white">产品规格管理</a></td>
    
  </tr>
</table>
<form method="POST" name="myform" onSubmit="return Checkly()">

<table  width="99%"  border="0" cellpadding="6" cellspacing="0"bgcolor="#f9f9f9" align="center">
<tr > 
 <td width="700"  >属性值名称：</br></br>
<input  name="s_value" type="text" id="s_value"  size="20"  /><font color="#FF0000">*</font></td>
</tr>

<tr > 
 <td  >排序：</br></br>
<input name="s_order" type="text" id="s_order"   value="1" style="width:40px"></td>
</tr>
<tr > 
<td height="30"  style="border-bottom:none"> <input class="bnt" type="submit" name="Submit" value="提交">
</td>
</tr>       
</table>
</form >

</body>
</html>
