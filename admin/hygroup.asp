<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="class_page.asp"-->
<%
if Instr(session("AdminPurview"),"|1000,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</script>"
  response.end
end if
'========判断是否具有管理权限
%>

<% 
If request.Form("button") = "批量删除" Then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='hygroup.asp';</script>" 
response.end()
end if
dim sql6 
sql6="delete from hygroup where id in ("&Request("id")&")"
conn.Execute ( sql6 )
addlog("删除会员组别")
end if
%>

<%if Request.QueryString("sc")="zm" then
dim sql,rs,ID
ID=request("ID")
set rs=server.createobject("adodb.recordset")
sql="delete from hygroup where Id="& ID
rs.open sql,conn,1,3
set rs=nothing
rs.close
 response.Write "<script language=javascript>alert(""会员组删除成功"");window.location.href=""hygroup.asp""</script>"
    response.End
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>zmcms_管理页面</title>
 <link href="css/style.css" rel="stylesheet" type="text/css">
<SCRIPT language=javascript>
function CheckAll(){ 
 for (var i=0;i<eval(zmqy.elements.length);i++){ 
  var e=zmqy.elements[i]; 
  if (e.name!="allbox") e.checked=zmqy.allbox.checked; 
 } 
} 
function ConfirmDel()
{
   if(confirm("确定要删除选中会员组吗？一旦删除将不能恢复！"))
     return true;
   else
     return false;
	 
}

</SCRIPT>
<script language="JavaScript" src="js/jquery-1.7.2.min.js"></script>
<style type="text/css">
.table1{ font-size: 14px; background: #F9f9f9;border:none;}
.table1 td{padding:10px; font-family:"微软雅黑";color:#ccc; font-weight:bold;}
.table1 td a{color:#ffffff}
</style>
<!--#include file="hqyy.asp"-->

</head>

<body>


<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">会员组管理 <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>


<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="huiyuan.asp" target="main">返回菜单</a></div><div class="dmeun"><a href="hygroupadd.asp" target="main">添加会员组</a></div></td>
 

 </tr>
</table>
<form name="zmqy" method="Post">

<%
DB_Sql = "select * from hygroup where lang='"&lang&"' "
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

Obj_Page.Setpage_Pathurl = "hygroup.asp" '执行分页的页面
Obj_Page.Setpage_Navtype = 1 '分页导航样式(可选1文字,2数字)

Call Obj_Page.Page_Init() '初始化分页
%>
              <table width="99%" align="center" cellpadding="0" cellspacing="0"  class="dtable2">
                
                <tr  class="wz">
                  <td width="3%"   ><div align="center">全选</div></td>
                  <td width="3%"   ><div align="center">ID</div></td> 
                  <td width="15%"  ><div align="center">会员组别号</div></td>
                  <td width="12%"   ><div align="center">会员组别名称</div></td>
                  <td width="10%"  ><div align="center">权限值</div></td>
                  <td width="12%"  ><div align="center">说明</div></td>
                  <td width="12%"  ><div align="center">创建时间</div></td>
                  <td width="8%"   > 
                    <div align="center">操作</div></td>
                  <td width="19%"   > 
                  <div align="center" style="border-right:none">删除</div></td>
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
  <%if rs("id")=1 or  rs("id")=2   then%><font color="red">×</font><%else%><input name="ID" type="checkbox" id="ID" value="<%=rs("id")%>" class="none"/><%end if%>
  
  
 </div></div></td>
                  <td  ><div align="center"><%=rs("id")%></div></td> 
                  <td ><div align="center"><%=rs("groupid")%></div></td>
                  <td ><div align="center"><%=rs("groupname")%></div></td>
                  <td ><div align="center"><%=rs("GroupLevel")%></div></td>
                  <td ><div align="center"><%=rs("Explain")%></div></td>
                  <td ><div align="center"><%=rs("adddate")%></div></td>
                  <td height="40" > 
<div align="center">
					  
					  
                        <%response.write "<a href='hygroupedit.asp?ID="&rs("Id")&"'>修改</a>"%>                  
                      </div></td>
                  <td > 
                   
					<div align="center"><%if rs("id")=1 or rs("id")=2 then%><font color="#FF0000">内置会员组无法删除</font><%else%><a href="hygroup.asp?id=<%=rs("id")%>&sc=zm" onClick="return ConfirmDel();">删除</a><%end if%> </div></td>
                </tr>
 <%
rs.MoveNext
Next	
%>   
<tr><td ><div align="center">
  <input type="checkbox" name="allbox" onClick="CheckAll()" class="none"/>
</div></td><td colspan="8" style="padding:5px; text-align:left"><input class="bnt" type="submit" name="button" id="button" value="批量删除"   onClick="return confirm('确定要批量删除选中的会员组吗，将不可恢复?');"/>
			  
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
