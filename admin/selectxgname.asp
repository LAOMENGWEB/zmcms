<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="Inc/Function.asp"-->
<%
if Instr(session("AdminPurview"),"|36,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</script>"
  response.end
end if
'========判断是否具有管理权限
%>
<%
mode=request("mode")
s_type = checknum(request("s_type"),0)
typeid=request("typeid")
s_order = checknum(request("s_order"),0)
id = checknum(request("id"),0)
If request.Form("submit") = "修改" Then
sql = "select * from [SelectName] where s_id = "&id&""
set rs = server.CreateObject("adodb.recordset")
rs.Open sql, conn, 1, 3

rs("s_name")=safereplace(request("s_name"))
rs("s_order")=s_order
rs("s_mode")=mode
rs("s_type")=s_type
rs("typeid")=typeid
rs("lang")=lang
rs.Update
rs.Close : set rs=Nothing
sql20="update SelectValue set typeid="&typeid&" where s_id="&id&""
conn.Execute ( sql20 )		
call addlog("修改属性名称")
response.Write "<script language=javascript>alert(""恭喜您给修改成功"");window.location.href=""select.asp?mode="&mode&"""</script>"
end if


Set rs2 = server.CreateObject("adodb.recordset")
sql2 = "select * from SelectName where s_id = "&id&" and lang='"&lang&"' "
rs2.Open sql2, conn, 1, 1
If rs2.EOF Then
    response.Write "<script language=javascript>window.location.href=""shuxing.asp""</script>"
    response.End
End If




%>



<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>附属属性类别增加</title>

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





<form method="POST" name="myform" class="checkform">
<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center"> <%if mode=0 then%> 修改筛选属性<%end if%>
     <%if mode=1 then%>修改规格属性<%end if%>
     <%if mode=2 then%>修改商品属性<%end if%>  <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>


<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="shuxing.asp" target="main">返回菜单</a></div>
<%if mode=0 then%> 
<div class="dmeun"><a href="select.asp?mode=0" target="main">筛选属性管理</a></div>
<%end if%>
<%if mode=1 then%> 
<div class="dmeun"><a href="select.asp?mode=1" target="main">规格属性管理</a></div>
<%end if%>
<%if mode=2 then%> 
<div class="dmeun"><a href="select.asp?mode=2" target="main">商品属性管理</a></div>
<%end if%>

 </td>
 

 </tr>
</table>


<table  width="99%"  border="0" cellpadding="6" cellspacing="0"  class="dtable4" align="center">
<tr > 
  <td width="11%"   >属性名称：</td>
  <td width="89%">
    <input  name="s_name" type="text" id="s_name" value="<%=rs2("s_name")%>"  size="20"  datatype="*" nullmsg="属性名称不能为空"/> <font color="#FF0000">*</font></td>
</tr>
 <%if mode=2 then%>
    <tr >
      <td  >属性类型：</td>
  <td>
      
      <input type="radio" name="s_type" value="0" class="none"  
   <%if rs2("s_type")="0" then response.write "checked"%>   
      
      
      
      > 单选
<input type="radio" name="s_type" value="2" class="none"
<%if rs2("s_type")="2" then response.write "checked"%>


> 多选
<input type="radio" name="s_type" value="1" class="none"

<%if rs2("s_type")="1" then response.write "checked"%>


> 文本域
      
      
      <%end if%>
      </td>
    </tr>
 <tr id="s_valuebox" style=" display:none"> 
   <td>属性值：</td>
  <td>
   
   <textarea id="s_value" name="s_value" cols="30" rows="20" class="text"></textarea>
   
   
   
   
   
   </td>
</tr>



		  
    <tr >
      <td  >所属类别：</td>
  <td>
      
      
      
      
 <select name="typeid">
<%Set rs = conn.execute("Select s_id,s_name from [SelectType] where lang='"&lang&"' order by s_order desc")
Do While Not rs.eof 
if rs2("typeid")=rs(0) then
Response.Write("<option value="""& rs(0) &"""  selected>"& rs(1) &"</option>")
else
Response.Write("<option value="""& rs(0) &""" >"& rs(1) &"</option>")
end if
rs.movenext 
Loop
rs.close : Set rs = Nothing 
%>
</select>
  </td>
    </tr>
   
    <tr > 
      <td  >排序：</td>
  <td>
      <input name="s_order" type="text" id="s_order"   value="<%=rs2("s_order")%>" style="width:40px"/> </td>
      </tr>
          
          <tr > 
            <td  colspan="2"> <input class="bnt" type="submit" name="Submit" value="修改">
            　            </td>
          </tr>       
</table>
</form >

</body>
</html>
