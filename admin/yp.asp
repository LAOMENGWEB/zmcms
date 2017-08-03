<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="class_page.asp"-->
<!--#include file="htmlconfig.asp"-->
<!--#include file="admin_html_function.asp"-->
<%
if Instr(session("AdminPurview"),"|73,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</script>"
  response.end
end if
'========判断是否具有管理权限
%>
<%
sh=request.QueryString("sh")
pone=request.QueryString("ypid")
If request.Form("button") = "批量删除" Then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='yp.asp';</script>" 
response.end()
end if
dim sql6 
sql6="delete from HrDemandAccept where id in ("&Request("id")&")"
conn.Execute ( sql6 )
call addlog("删除应聘信息")
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
<SCRIPT language=javascript>
function CheckAll(){ 
 for (var i=0;i<eval(zmqy.elements.length);i++){ 
  var e=zmqy.elements[i]; 
  if (e.name!="allbox") e.checked=zmqy.allbox.checked; 
 } 
} 
function ConfirmDel()
{
   if(confirm("确定要删除选中的应聘吗？一旦删除将不能恢复！"))
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
<!--#include file="hqyy.asp"-->
</head>

<body>


<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">应聘管理  <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>


<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td >
 <div class="dmeun"><a href="zhaopin.asp" target="main">返回菜单</a></div>
 <div class="dmeun"><a href="yp.asp" target="main">所有应聘</a></div>
 <div class="dmeun"><a href="yp.asp?sh=2" target="main">已回复应聘</a></div>
 <div class="dmeun"><a href="yp.asp?sh=3" target="main">未回复应聘</a></div>
 
 </td>
 

 </tr>
</table>


<form name="zmqy" method="Post">
<%


if pone="" then
DB_Sql = "select * from HrDemandAccept where lang='"&lang&"' "
end if
if pone<>"" then
DB_Sql = "select * from HrDemandAccept where MemID="&pone&" and lang='"&lang&"' "
end if
if sh=2 then
DB_Sql = "select * from HrDemandAccept where ReplyContent is not null and lang='"&lang&"' "
end if
if sh=3 then
DB_Sql = "select * from HrDemandAccept where ReplyContent is null and lang='"&lang&"' "
end if

DB_Sql = DB_Sql & "order by ID desc"
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
if pone="" then
Obj_Page.Setpage_Pathurl = "yp.asp" '执行分页的页面
end if

if pone<>"" then
Obj_Page.Setpage_Pathurl = "yp.asp?ypID="&pone&"" '执行分页的页面
end if

if sh=2 then
Obj_Page.Setpage_Pathurl = "yp.asp?sh=2" '执行分页的页面
end if
if sh=3 then
Obj_Page.Setpage_Pathurl = "yp.asp?sh=3" '执行分页的页面
end if
Obj_Page.Setpage_Navtype = 1 '分页导航样式(可选1文字,2数字)

Call Obj_Page.Page_Init() '初始化分页

%>

<table width="99%" border="0" cellpadding="0" cellspacing="0"  align="center" class="dtable2">
  <tr class="wz">
    <td width="5%"  ><div align="center">全选</div></td>
    <td width="3%"  ><div align="center">编号</div></td>
    <td width="11%"  ><div align="center">应聘岗位</div></td>
    <td width="10%"  ><div align="center">姓名</div></td>
    <td width="9%"  ><div align="center">联系电话</div></td>
    <td width="15%"  ><div align="center">电子邮件</div></td>
    <td width="12%"  ><div align="center">应聘日期</div></td>
    <td width="11%"  ><div align="center">回复日期</div></td>
    <td width="12%"  ><div align="center">查看详情</div></td>
    <td width="12%" ><div align="center">操作选性</div></td>
  </tr>
  <%

'/循环开始
For II = 1 to Obj_PageListnum
If rs.Eof Then
	Exit For
End If
%>
  <tr>
    <td ><div align="center">
      <input name="ID" type="checkbox" id="ID" value="<%=rs("id")%>" style="border:none"/>
    </div></td>
    <td><div align="center"><%=rs("ID")%></div></td>
    <td ><div align="center"><%=rs("Quarters")%></div></td>
    <td><div align="center"><%=Guest(rs("MemID"),rs("name"))%></div></td>
    <td><div align="center"><%=rs("phone")%></div></td>
    <td><div align="center"><%=rs("email")%><a href="sendemail.asp?l=somail&selectemail=<%=rs("email")%>"><font color="#FF0000">回信</font></a></div></td>
    <td><div align="center"><%=rs("Adddate")%></div></td>
    <td><div align="center">
      <%
if trim(rs("replycontent")&"")="" then
response.Write("<font color=#FF0000>此应聘暂未回复！</font>")
else
%>
          <%=rs("ReplyTime")%>
          <%
end if
  %>
    </div></td>
    <td><div align="center"><a  href="ypxx.asp?id=<%=rs("id")%>">查看详请</a></div></td>
    <td ><div align="center"><a href="delyp.asp?id=<%=rs("ID")%>" onClick="return ConfirmDel();">删除</a> </div></td>
  </tr>
    <%
rs.MoveNext
Next
	%> 
	<tr>
  <td><div align="center">
    <input type="checkbox" name="allbox" onClick="CheckAll()" style="border:none"/>
  </div></td>
  <td colspan="11" style="padding:5px; text-align:left">  <input class="bnt" type="submit" name="button" id="button" value="批量删除"   onClick="return confirm('确定要批量删除选中的应聘吗，将不可恢复?');"/></td>  
  </tr>
</table>

            
    <%If Obj_Page.OutErrInfo <> "" Then
		Response.Write Obj_Page.OutErrInfo '报错信息显示
		Response.End
		else
	Call Obj_Page.Page_Navigation() '分页导航
	end if
    else
	%>
<table width="99%" border="0" align="center" cellpadding="0" cellspacing="0" class="table1" style="margin-top:10px">
  <tr>
    <td style="border:right:none;border-bottom:none">	<div align="center">
      <%
Response.Write "目前暂无任何信息"
Response.end
	%>
    </div></td>
  </tr>
</table>

	<%Set Obj_Page = Nothing

RS.Close
Set RS = Nothing
End If
 conn.close
 set conn=nothing
%>
</form>
</body>
</html>
<%
function Guest(ID,Linkman)
  Dim rs,sql
  Set rs=server.CreateObject("adodb.recordset")
  sql="Select * From Members where ID="&ID&" and lang='"&lang&"'"
  rs.open sql,conn,1,1
  if rs.bof and rs.eof then
    Guest="{未注册用户}"&Linkman&""
  else
    Guest="<font color='red'>{"&rs("GroupName")&"}</font><a href='Memberxg.asp?Result=Modify&ID="&ID&"'>"&Linkman&"</a>"
  end if
  rs.close
  set rs=nothing
end function
%>