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
id=request.QueryString("id")
If request.Form("submit") = "修改" Then
s_name=Trim(request.form("s_name"))
zmone=trim(request.form("zmone"))
zmtwo=trim(request.form("zmtwo"))
zmthree=trim(request.form("zmthree"))
s_order=trim(request.form("s_order"))
ClassId=trim(Request.Form("ClassId"))
	  Set rs=server.CreateObject("adodb.recordset")
	  sql="Select * From cpClass where ID="&ClassId  
	  rs.open sql,conn,1,1
	  if Not rs.bof and Not rs.eof then
	  ParentClassID=rs("ParentClassID")  
	  end if
	  rs.close
	  if ParentClassID<>0 then
	  Set rs=server.CreateObject("adodb.recordset")
	  sql="Select * From cpclass where id="&ParentClassID  
	  rs.open sql,conn,1,1
	  if Not rs.bof and Not rs.eof then  
	  ParentClassID1=rs("ParentClassID")  
	  end if
	  rs.close
	  end if
      if ParentClassID1<>0 then
	  Set rs=server.CreateObject("adodb.recordset")
	  sql="Select * From cpclass where id="&ParentClassID1 
	  rs.open sql,conn,1,1
	  if Not rs.bof and Not rs.eof then  
	  ParentClassID2=rs("ParentClassID")     
	  end if
	  rs.close
	  end if
	  
Set rs = server.CreateObject("adodb.recordset")
sql1="Select * From SelectType Where s_id="&id
rs.open sql1,conn,1,3

	  
	 if ParentClassID=0 then
rs("zmone")=trim(Request.Form("ClassId"))
rs("zmtwo")=0
rs("zmthree")=0
end if
if ParentClassID<>0 and ParentClassID1=0 then
rs("zmone")=ParentClassID
rs("zmtwo")=trim(Request.Form("ClassId"))
rs("zmthree")=0
end if
if ParentClassID<>0 and ParentClassID1<>0 and ParentClassID2=0 then
rs("zmone")=ParentClassID1
rs("zmtwo")=ParentClassID
rs("zmthree")=trim(Request.Form("ClassId"))
end if

rs("s_name")=s_name
rs("s_order")=s_order
rs("lang")=lang
rs.update
rs.close
set rs=nothing 
call addlog("修改附属属性类别")

response.Write "<script language=javascript>alert(""恭喜您修改成功"");window.location.href=""selectvaule.asp""</script>"


end if
Set rs = server.CreateObject("adodb.recordset")
sql2 = "select * from SelectType where s_id = "&request.QueryString("id")&" and Lang='"&Language&"'"
rs.Open sql2, conn, 1, 1
If rs.EOF Then
    response.Write "<script language=javascript>window.location.href=""selectvaule.asp""</script>"
    response.End
End If

if rs("zmone")<>0 and  rs("zmtwo")=0  and rs("zmthree")=0 then
ClassId=rs("zmone")
end if
if rs("zmone")<>0 and  rs("zmtwo")<>0 and rs("zmthree")=0  then
ClassId=rs("zmtwo")
end if
if rs("zmone")<>0 and  rs("zmtwo")<>0 and rs("zmthree")<>0  then
ClassId=rs("zmthree")
end if



%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>属性名称修改</title>
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
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">修改属性类别  <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>




<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="shuxing.asp" target="main">返回菜单</a></div>
<div class="dmeun"><a href="selectvaule.asp" target="main">属性类别管理</a></div>
 </td>
 

 </tr>
</table>
<table  width="99%"  border="0" cellpadding="0" cellspacing="0"  class="dtable4" align="center">
<tr > 
  <td width="11%"  >属性类别名称：</td>
	  <td width="89%">
    <input  name="s_name" type="text" id="s_name"  size="20"  value="<%=rs("s_name")%>" datatype="*" nullmsg="类别名称不能为空"/> <font color="#FF0000">*</font></td>
</tr>

 <tr > 
   <td >所属类别：</td>
	  <td><select ID="ClassID" name="ClassID" datatype="*" nullmsg="所属类别不能为空">
     <%call mxlb("cpClass",ClassId)%>
     </select> <font color="#FF0000">*</font><font color=red>（有三级栏目的必须选择三级栏目）</font></td>
</tr>



		  
    <tr > 
      <td  >排序：</td>
	  <td>
      <input name="s_order" type="text" id="s_order"   value="<%=rs("s_order")%>" style="width:40px"></td>
      </tr>
          
          <tr > 
            <td colspan="2"> <input class="bnt" type="submit" name="Submit" value="修改">
            　            </td>
          </tr>       
</table>
</form >

</body>
</html>
