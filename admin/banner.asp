<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="class_page.asp"-->
<%
if Instr(session("AdminPurview"),"|111,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</script>"
  response.end
end if
'========判断是否具有管理权限
%>

<%
pone=request.QueryString("pone")
ptwo=request.QueryString("ptwo")
pthree=request.QueryString("pthree")
If request.Form("button") = "批量删除" Then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='banner.asp';</script>" 
response.end()
end if

dim id
ID=trim(request("ID"))
ID1=split(ID,",")
for i=0 to ubound(ID1)
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from banner where ID in ("&id1(i) &")",conn,0,1
call DeleteFile1(rs("uploadfile"))
next
conn.Execute ( "delete from banner where id in ("&Request("id")&")" )
call addlog("删除幻灯片")
response.Write "<script language=javascript>alert(""删除成功！"");window.location.href=""banner.asp""</script>"
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
   if(confirm("确定要删除选中的幻灯图片吗？一旦删除将不能恢复！"))
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


<form name="zmqy" method="Post">


<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">幻灯列表管理  <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>


<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="huandeng.asp" target="main">返回菜单</a></div><div class="dmeun"><a href="banneradd.asp" target="main">添加幻灯</a></div><div class="dmeun"><a href="fl_banner.asp" target="main">幻灯分类管理</a></div></td>
 

 </tr>
</table>


<table width="99%" align="center" cellpadding="0" cellspacing="0" class="dtable1">
  		
          <tr   >
            <td  >按照分类查找：<select name="ClassID" id="ClassID" onChange="javascript:window.open(this.options[this.selectedIndex].value,'main')">
    <%
   Dim Sqlp,Rsp,TempStr
   Sqlp ="select * from bannerclass where ParentClassID=0 order by px asc"   
   Set Rsp=server.CreateObject("adodb.recordset")   
   rsp.open sqlp,conn,1,1 
   Response.Write("<option value="""">请选择分类</option>") 
   If Rsp.Eof and Rsp.Bof Then
      Response.Write("<option value="""">请先添加分类</option>")
   Else
	   Response.Write("<option value=""banner.asp"">所有幻灯信息</option>") 
      Do while not Rsp.Eof   
         Response.Write("<option value=" & """?pone=" & Rsp("ID") & """" & "")
		 If int(request("pone"))=Rsp("ID") then
				Response.Write(" selected" ) 
		 End if
         Response.Write(">" & Rsp("ClassName") & " </option>")  & VbCrLf
		 
      Rsp.Movenext   
      Loop   
   End if
	rsp.close
	set rsp=nothing
	%>
        </select></td>
 
  </tr>
</table>


<%
if pone="" then
DB_Sql = "select * from banner where lang='"&lang&"' "
end if
if pone<>""  then
DB_Sql = "select * from banner where zmone="&Pone&" and lang='"&lang&"' "
end if
DB_Sql = DB_Sql & "order by banner_order asc"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open DB_Sql,Conn,1,1
If Not(rs.Bof and rs.Eof) Then
Dim Obj_Page, Obj_PageListnum, II
Obj_PageListnum = 20 

Set Obj_Page = New HHYY_Pageclass

Obj_Page.Setpage_RS = rs
Obj_Page.Setpage_Listnum = Obj_PageListnum
if pone="" then
Obj_Page.Setpage_Pathurl = "banner.asp" '执行分页的页面
end if

if pone<>"" then
Obj_Page.Setpage_Pathurl = "banner.asp?pone="&pone&"" '执行分页的页面
end if
Obj_Page.Setpage_Navtype = 1 '分页导航样式(可选1文字,2数字)

Call Obj_Page.Page_Init() '初始化分页

%>

<table width="99%"   align="center" cellpadding="0" cellspacing="0" class="dtable2">

        <tr>
          <td width="47" align="center" bgcolor="#eeeeee" >全选</td>
          <td width="47" align="center" bgcolor="#eeeeee">编号</td>
          <td width="47" height="30" align="center" bgcolor="#eeeeee"><div align="center">排序</div></td>
          <td width="237" bgcolor="#eeeeee"><div align="center">所属分类</div></td>
          <td width="237" bgcolor="#eeeeee"><div align="center" class="STYLE2">图片预览</div></td>
          <td width="280" bgcolor="#eeeeee" ><div align="center" class="STYLE2">幻灯名称</div></td>
          <td width="244" bgcolor="#eeeeee"><div align="center" class="STYLE2">链接地址</div></td>
          <td width="255" bgcolor="#eeeeee" style="border-right:none"><div align="center" class="STYLE2">修改操作</div></td>
        </tr>


 <%

'/循环开始
For II = 1 to Obj_PageListnum
If rs.Eof Then
	Exit For
End If
%>

        <tr align="center">
          <td ><div align="center">
            <input name="ID" type="checkbox" id="ID" value="<%=rs("id")%>" />
          </div></td>
          <td ><%=rs("id")%></td>
          <td ><iframe name="zmcms" width="150" height="50" frameborder="0" scrolling="no" 
				src="xghdpx.asp?id=<%=rs("id")%>" ></iframe></td>
          <td ><%
		  dim rsc
		  set rsc=server.CreateObject("adodb.recordset")
		  rsc.open "select * from bannerclass",conn,1,1
		  while not rsc.eof
		    if rs("zmone")=rsc("id") then
			response.Write("" & rsc("classname") & "")
			end if
			rsc.movenext
		wend
		rsc.close
		set rsc=nothing
		  %></td>
          <td ><img src="../<%=rs("uploadfile")%>" width="200" height="40"  style="padding:10px 0"/></td>
          <td ><%=rs("banner_name")%></td>
          <td ><%=rs("banner_url")%></td>
          <td style="border-right:none">[<a href="banner_modify.asp?id=<%=rs("ID")%>">修改</a>] [<a href="banner_delete.asp?id=<%=rs("ID")%>" onClick="return confirm('确定要删除此幻灯吗？')">删除</a>]</td>
        </tr>
  <%
rs.MoveNext
Next
	
%> 
<tr>
  <td><div align="center">
    <input type="checkbox" name="allbox" onClick="CheckAll()" />
  </div></td>
  <td colspan="10" style="padding:5px;text-align:left">  <input class="bnt" type="submit" name="button" id="button" value="批量删除"   onClick="return confirm('确定要批量删除选中的幻灯片吗，将不可恢复?');"/></td>  
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
	<table width="99%" border="0" align="center" cellpadding="0" cellspacing="0" class="table1">
  <tr>
    <td style="padding:10px 0 ;border-right:none;border-bottom:none">	<div align="center">
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
%>
</form>
</body>
</html>