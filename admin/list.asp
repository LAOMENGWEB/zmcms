<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="class_page.asp"-->
<%
url=Request.ServerVariables("HTTP_REFERER")
If request.Form("B1") = "批量删除" Then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='"&url&"';</script>" 
response.end()
end if
for each i in request.form("id")
if i<>"" then
conn.execute "delete from Records where cid="&i&""
conn.execute "delete from info where id="&i&"" 
end if
next
call addlog("批量删除表单信息")
end if

If request.Form("B1") = "标记已读" Then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='"&url&"';</script>" 
response.end()
end if
dim sql4 
sql4="update info set bj=1 where id in ("&Request("id")&")"
conn.Execute ( sql4 )
call addlog("处理表单信息标记")
end if
If request.Form("B1") = "取消标记" Then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='"&url&"';</script>" 
response.end()
end if
dim sql5 
sql5="update info set bj=0 where id in ("&Request("id")&")"
conn.Execute ( sql5 )
call addlog("取消表单信息标记")
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

		<script>
function CheckAll(form)
  {
  for (var i=0;i<form.elements.length;i++)
    {
    var e = form.elements[i];
    if (e.name != 'chkall')
       e.checked = form.chkall.checked;
    }
  }
                                          </script>
</head>

<body>
 <table width="100%" border="0" cellpadding="0" cellspacing="0"   class="dtable">
<tr>
 <td >查看反馈</td>
</tr>
</table>




<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="chajian.asp" target="main">返回菜单</a></div><div class="dmeun"><a href="diyformmanage.asp" target="main">自定义表单管理</a></div></td>
 

 </tr>
</table>
<form method="POST" name="form">

<%
pid=request.QueryString("pid")

DB_Sql = "select * from info where cc=1 "
DB_Sql = DB_Sql & "order by  ID desc"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open DB_Sql,Conn,1,1
If Not(rs.Bof and rs.Eof) Then
Dim Obj_Page, Obj_PageListnum, II
Obj_PageListnum = 10
Set Obj_Page = New HHYY_Pageclass
Obj_Page.Setpage_RS = rs
Obj_Page.Setpage_Listnum = Obj_PageListnum
Obj_Page.Setpage_Pathurl = "list.asp?pid="&pid&"" '执行分页的页面
Obj_Page.Setpage_Navtype = 1 '分页导航样式(可选1文字,2数字)

Call Obj_Page.Page_Init() '初始化分页

%>
<TABLE width="99%"  border="0" align="center" cellpadding="0" cellspacing="0" class="dtable2">
<TBODY>
<TR class="wz">
  <TD  ><div align="center">全选</div></TD>
  <TD ><div align="center">ID</div></TD>
  <%
  set rs2=server.createobject("adodb.recordset")
sql2="select top 5  * from Records where bid='"&rs("name")&"'and pid="&pid&" order by id asc"
rs2.open sql2,conn,1,3
do while not rs2.eof

%>
  
  <TD    ><div align="center"><%=rs2("name")%></div></TD>
  <%
rs2.movenext
loop
rs2.close
set rs2=nothing
%>
  
  <TD   ><div align="center">提交日期</div></TD>
  <TD   ><div align="center">是否已读</div></TD>
  <TD   style="border-right:none"><div align="center">查看详情</div></TD>
</tr>
<%
'/循环开始
For II = 1 to Obj_PageListnum
If rs.Eof Then
	Exit For
End If

set rs1=server.createobject("adodb.recordset")
sql="select top 5 * from Records where bid='"&rs("name")&"'and pid="&pid&" "
rs1.open sql,conn,1,3

%>

<TR>
<TD ><div align="center"><input type="checkbox" name="id" value="<%=rs("id")%>" class="none">
  </div></TD>

<TD><div align="center"><%=rs("id")%></div></TD>

 <%
  set rs2=server.createobject("adodb.recordset")
sql2="select top 5 * from Records where bid='"&rs("name")&"'and pid="&pid&" order by id asc"
rs2.open sql2,conn,1,3
do while not rs2.eof

%>
  
  <TD  ><div align="center"><%=rs2("Record")%></div></TD>
  <%
rs2.movenext
loop
rs2.close
set rs2=nothing
%>





<TD ><div align="center"><%=rs("riqi")%></div></TD>
<TD >
  <div align="center">
    <%
if rs("bj")=1 then
response.Write("<font color=#FF0000>已处理</font>")
else
response.Write("<font color=#FF0000>未处理</font>")
end if


%>
  </div></TD>
<TD  >
  <div align="center"><a href="show.asp?bid=<%=rs("name")%>&id=<%=rs("id")%>&pid=<%=rs1("pid")%>">查看详情</a></div></TD>
</tr>
<%

rs1.close
set rs1=nothing

rs.MoveNext
Next	
%>      

<TR>
  <TD   > 
 
      <div align="center">
        <input type="checkbox" name="chkall" value="on" onClick="CheckAll(this.form)" class="none">
      </div></td>
	  <td colspan="10" style=" padding:5px; text-align:left">
	  <input type="submit" value="批量删除" name="B1" class="bnt" />
	  <input type="submit" value="标记已读" name="B1" class="bnt" />
	  <input type="submit" value="取消标记" name="B1" class="bnt" />
	  </TD>
  </tr>
</TBODY></TABLE>
<%
If Obj_Page.OutErrInfo <> "" Then
Response.Write Obj_Page.OutErrInfo '报错信息显示
Response.End
else
Call Obj_Page.Page_Navigation() '分页导航
end if
else
%>

<table width="99%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td style="padding:10px 0 ;border-right:none;border-bottom:none">	<div align="center">
      <%
Response.Write "目前暂无任何信息"
Response.end
	%>
    </div></td>
  </tr>
</table>


<%
Set Obj_Page = Nothing
RS.Close
Set RS = Nothing
End If
%>
 
 
</form>


</body>
</html>
