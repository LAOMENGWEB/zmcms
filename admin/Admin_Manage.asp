<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="class_page.asp"-->
<%
if Instr(session("AdminPurview"),"|01,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</script>"
  response.end
end if
'========判断是否具有管理权限
%>

<% 
If request.Form("button") = "设为生效" Then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='admin_manage.asp';</script>" 
response.end()
end if
dim sql4 
sql4="update admin set working=1 where id in ("&Request("id")&")"
call addlog("启用管理员账号")
conn.Execute ( sql4 )
end if
If request.Form("button") = "取消生效" Then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='admin_manage.asp';</script>" 
response.end()
end if
dim sql5 
sql5="update admin set working=0 where id in ("&Request("id")&")"
conn.Execute ( sql5 )
call addlog("禁用管理员账号")
end if
If request.Form("button") = "批量删除" Then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='admin_manage.asp';</script>" 
response.end()
end if
dim sql6 
sql6="delete from admin where id in ("&Request("id")&")"
conn.Execute ( sql6 )
call addlog("删除管理员账号")
end if
%>

<%if Request.QueryString("sc")="zm" then
dim sql,rs,ID
ID=request("ID")
set rs=server.createobject("adodb.recordset")
sql="delete from admin where Id="& ID
rs.open sql,conn,1,3
set rs=nothing
call addlog("删除管理员账号")
 response.Write "<script language=javascript>alert(""管理员删除成功"");window.location.href=""admin_manage.asp""</script>"
    response.End
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>zmcms_管理页面</title>
<script language="JavaScript" src="../js/jquery-1.7.2.min.js"></script>
 <script type="text/javascript" src="../js/Validform_v5.3.2_min.js"></script>
 <link href="css/style.css" rel="stylesheet" type="text/css">
  <script type="text/javascript">
$(function(){
	
$(".checkform").Validform();  //就这一行代码！;
	
})
</script>
<SCRIPT language=javascript>
function CheckAll(){ 
 for (var i=0;i<eval(zmqy.elements.length);i++){ 
  var e=zmqy.elements[i]; 
  if (e.name!="allbox") e.checked=zmqy.allbox.checked; 
 } 
} 
function ConfirmDel()
{
   if(confirm("确定要删除选中的管理员吗？一旦删除将不能恢复！"))
     return true;
   else
     return false;
	 
}

</SCRIPT>

<style type="text/css">
.table1{ font-size: 14px; background: #F9f9f9;border:none;}
.table1 td{padding:10px; font-family:"微软雅黑";color:#ccc; font-weight:bold;}
.table1 td a{color:#ffffff}
</style>
</head>

<body>
<form name="zmqy" method="Post">


<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">管理员账号管理 </div></td>
</tr>
</table>


<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="zhanghao.asp" target="main">返回菜单</a></div><div class="dmeun"><a href="admin_add.asp" target="main">添加管理员账号</a></div></td>
 

 </tr>
</table>
<%
DB_Sql = "select * from admin "
DB_Sql = DB_Sql & "order by ID asc"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open DB_Sql,Conn,1,1
If Not(rs.Bof and rs.Eof) Then


'================================================
'调用分页类开始
Dim Obj_Page, Obj_PageListnum, II
Obj_PageListnum = 20 

Set Obj_Page = New HHYY_Pageclass

Obj_Page.Setpage_RS = rs
Obj_Page.Setpage_Listnum = Obj_PageListnum

Obj_Page.Setpage_Pathurl = "admin_manage.asp" '执行分页的页面
Obj_Page.Setpage_Navtype = 1 '分页导航样式(可选1文字,2数字)

Call Obj_Page.Page_Init() '初始化分页
%>
              <table width="99%" align="center" cellpadding="0" cellspacing="0"  class="dtable2">
                
                <tr  >
                  <td width="3%" ><div align="center">全选</div></td>
                  <td width="3%" ><div align="center">ID</div></td> 
                  <td width="6%" > <div align="center">管理员帐号</div></td>
                  <td width="7%" ><div align="center">管理权限</div></td>
                  <td width="5%" ><div align="center">是否生效</div></td>
                  <td width="15%" ><div align="center">操作权限</div></td>
                  <td width="10%" ><div align="center">最后登录IP</div></td>
                  <td width="12%" ><div align="center">最后登录时间</div></td>
                  <td width="12%" ><div align="center">创建时间</div></td>
                  <td width="8%" > 
                  <div align="center">操作</div></td>
                  <td width="19%" style="border-right:none"> 
                  <div align="center">删除</div></td>
                </tr>
<%
'/循环开始
For II = 1 to Obj_PageListnum
If rs.Eof Then
	Exit For
End If
%>
                <tr  >
                  <td  ><div align="center"><div align="center">
  <%if rs("id")<>1 then%><input name="ID" type="checkbox" id="ID" value="<%=rs("id")%>" class="none"/><%else%><font color="red">×</font><%end if%>
 </div></div></td>
                  <td  ><div align="center"><%=rs("id")%></div></td> 
                  <td   > 
                    <div align="center"><%=rs("UserName")%></div></td>
                  <td ><div align="center"><%=rs("glyName")%></div></td>
                  <td  > 
				    <div align="center">
				      <%
                  if rs("working")=1 then
        Response.Write "生效" 
      else
        Response.Write "未生效" 
	  end If%>
		          </div></td>
                  <td ><%if rs("id")=1 then%>超级管理员拥有所有的权限<%else%><%=left(rs("AdminPurview"),25)%><%end if%></td>
                  <td ><%if rs("lastloginip")<>"" then%><%=rs("lastloginip")%><%else%>此账号从未登陆<%end if%></td>
                  <td ><%if rs("lastlogintime")<>"" then%><%=rs("lastlogintime")%><%else%>此账号从未登陆<%end if%></td>
                  <td ><div align="center"><%=rs("adddate")%></div></td>
                  <td  > 
                    
                      <div align="center">
                        <%response.write "<a href='Admin_ManageEdit.asp?ID="&rs("Id")&"'>修改</a>"%>                  
                      </div></td>
                  <td  style="border-right:none"> 
                   
					<div align="center"><%if rs("id")<>1 then%><a href="Admin_Manage.asp?id=<%=rs("id")%>&sc=zm" onClick="return ConfirmDel();">删除</a><%else%><font color="#FF0000">内置管理员无法删除</font><%end if%> </div></td>
                </tr>
 <%
rs.MoveNext
Next	
%>   
<tr><td><div align="center">
  <input type="checkbox" name="allbox" onClick="CheckAll()" class="none"/>
</div></td><td colspan="10" style="padding:5px; text-align:left"> <input class="bnt" type="submit" name="button" id="button" value="设为生效" onClick="return confirm('确定要把选中的管理员启用吗？');" />
              <input class="bnt" type="submit" name="button" id="button" value="取消生效"   onClick="return confirm('确定要把选中的管理员禁用吗？');"/>
			  <input class="bnt" type="submit" name="button" id="button" value="批量删除"   onClick="return confirm('确定要批量删除选中的管理员吗，将不可恢复?');"/>
			  
			  </td></tr>
</table>
              <%If Obj_Page.OutErrInfo <> "" Then
		Response.Write Obj_Page.OutErrInfo '报错信息显示
		Response.End
		else
	Call Obj_Page.Page_Navigation() '分页导航
	end if
    else
Response.Write "目前暂无任何信息"
Response.end
	%>
<%Set Obj_Page = Nothing

RS.Close
Set RS = Nothing
End If
%>
</form>
</body>
</html>
