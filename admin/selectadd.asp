<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="Inc/Function.asp"-->

<%
if Instr(session("AdminPurview"),"|36,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</script>"
  response.end
end if
mode=request("mode")
s_type = checknum(request("s_type"),0)
typeid=request("typeid")
s_order = checknum(request("s_order"),0)
id = checknum(request("id"),0)
If request.Form("submit") = "提交" Then
If id = 0 Then 
sql = "select * from [SelectName]"
set rs = server.CreateObject("adodb.recordset")
rs.Open sql, conn, 1, 3
rs.AddNew
rs("s_name")=safereplace(request("s_name"))
rs("s_order")=s_order
rs("s_mode")=mode
rs("s_type")=s_type
rs("typeid")=typeid
rs("lang")=lang
rs.Update
rs.Close : set rs=Nothing
 
Set rs = conn.execute("select s_id from [SelectName] where s_name = '"& safereplace(request("s_name")) &"' and s_mode = "& mode &" and typeid = "& typeid &" and lang='"&lang&"'")
				id = rs(0)
				rs.close : Set rs = Nothing 
				End If 
		Arr_SelectValue = Split(Replace(Trim(request("s_value")),"'","&#180;"),Chr(13)&Chr(10))
		For ii = 0 To UBound(Arr_SelectValue)
			set rs = server.CreateObject("adodb.recordset")
			rs.Open "select * from [SelectValue] where s_value = '"& Split(Replace(Trim(request("s_value")),"'","&#180;"),Chr(13)&Chr(10))(ii) &"' and s_mode = "& mode &" and typeid = "& typeid &"", conn, 1, 3
			If rs.eof Then 
					rs.AddNew
					rs("s_mode")=mode
					rs("typeid")=typeid
					
			End If 
			
					s_value = Replace(Trim(request("s_value")),"'","&#180;")
					If s_value = "" Then s_value = Replace(Trim(request("s_value")),"'","&#180;")
					rs("s_value")=Split(s_value,Chr(13)&Chr(10))(ii)
			rs("lang")=lang
			rs("s_order")=UBound(Arr_SelectValue)+1-ii
			rs("s_id")=id
			rs("s_pic")=s_pic
			rs.Update
			rs.Close : set rs=Nothing
		Next 
		













call addlog("添加附属属性类别")

response.Write "<script language=javascript>alert(""恭喜您添加成功"");window.location.href=""select.asp?mode="&mode&"""</script>"


end if

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

<SCRIPT language=javaScript>

 
 
function UType(n){

    var s_valuebox = document.getElementById("s_valuebox");
	if (n == 1){
		s_valuebox.style.display = "none";
	}
	if (n == 0){
		s_valuebox.style.display = "";
	}

} 
 
 
 
 
 
 
 
</SCRIPT>
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
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center"> <%if mode=0 then%> 添加筛选属性<%end if%>
     <%if mode=1 then%>添加规格属性<%end if%>
     <%if mode=2 then%>添加商品属性<%end if%>  <span id="yuyan" style="margin-left:20px"></span></div></td>
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

<table  width="99%"  border="0" cellpadding="0" cellspacing="0"  class="dtable4" align="center">
<tr > 
  <td width="11%"  >属性名称：</td>
    
	<td width="89%">
	<input  name="s_name" type="text" id="s_name"  size="20"  datatype="*" nullmsg="属性名称不能为空"/> <font color="#FF0000">*</font></td>
</tr>
 <%if mode=2 then%>
    <tr >
      <td  >属性类型：</td>
    
	<td>
      
      <input type="radio"  class="none" name="s_type" value="0" onClick="UType(0)" checked> 单选
<input type="radio" name="s_type" value="2" onClick="UType(0)" class="none"> 多选
<input type="radio" name="s_type" value="1" onClick="UType(1)" class="none"> 文本域
      
      
      <%end if%>
      </td>
    </tr>
 <tr id="s_valuebox"> 
   <td >属性值：</td>
    
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
		Response.Write("<option value="""& rs(0) &""">"& rs(1) &"</option>")
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
      <input name="s_order" type="text" id="s_order"   value="1" style="width:40px"></td>
      </tr>
          
          <tr > 
            <td  colspan="2"> <input class="bnt" type="submit" name="Submit" value="提交">
            　            </td>
          </tr>       
</table>
</form >

</body>
</html>
