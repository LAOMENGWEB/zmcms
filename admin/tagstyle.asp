<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="class_page.asp"-->
<%
if Instr(session("AdminPurview"),"|611,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</script>"
  response.end
end if
'========判断是否具有管理权限
%>

<%
t= request.QueryString("t")
Result=request.QueryString("Result")
keytag=request.Form("k1")
t1=request.Form("t1")
pone=request.QueryString("pone")
ptwo=request.QueryString("ptwo")
pthree=request.QueryString("pthree")
If request.Form("button") = "批量删除" Then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='tagstyle.asp';</script>" 
response.end()
end if

conn3.Execute ( "delete from templateTags where id in ("&Request("id")&")" )
call addlog("删除标签样式")
if t=1 then
response.Write "<script language=javascript>alert(""删除成功！"");window.location.href=""tagstyle.asp?t=1""</script>"
end if
if t=2 then
response.Write "<script language=javascript>alert(""删除成功！"");window.location.href=""tagstyle.asp?t=2""</script>"
end if
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
   if(confirm("确定要删除选中的标签样式吗？一旦删除将不能恢复！"))
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
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center"><%if t=1 then %> PC端样式管理 <%end if%>
 <%if t=2 then %> WAP端样式管理 <%end if%></div></td>
</tr>
</table>




<table width="99%" align="center" cellpadding="0" cellspacing="0"  class="dtable1" >
  		
          <tr   >
            
        <form id="form2" name="form2" method="post" action="tagstyle.asp?Result=ok&t=<%=t%>">
            <td width="8%" >
             快速搜索：</td>
            <td width="92%">
                <input name="k1" type="text" id="k1" />
                <input name="t1" type="hidden" id="t1" value="<%=t%>"/>
      
             <input class="bnt" type="submit" name="button" id="button" value="搜索" />
          </td></form>
  </tr>
       
</table>

<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="biaoqian.asp" target="main">返回菜单</a></div>
 <div class="dmeun"><a href="tagstyleadd.asp?t=1" target="main">添加PC端标签</a></div>
 
 <div class="dmeun"><a href="tagstyleadd.asp?t=2" target="main">添加WAP端标签</a></div>
 
 </td>

 </tr>
</table>
<form name="zmqy" method="Post">
<%
if Result="ok"   then
if keytag<>"" then
 DB_Sql = "select * from templateTags where tagname like '%"&keytag&"%' and fl="&t&" "
 else
  DB_Sql = "select * from templateTags where  fl="&t&" "
 end if
 
else
DB_Sql = "select * from templateTags where fl="&t&" "
end if
DB_Sql = DB_Sql & "order by id desc"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open DB_Sql,Conn3,1,1
If Not(rs.Bof and rs.Eof) Then
Dim Obj_Page, Obj_PageListnum, II
Obj_PageListnum = 30 

Set Obj_Page = New HHYY_Pageclass

Obj_Page.Setpage_RS = rs
Obj_Page.Setpage_Listnum = Obj_PageListnum
if Result="ok"   then
Obj_Page.Setpage_Pathurl = "tagstyle.asp?Result=ok" '执行分页的页面
else
Obj_Page.Setpage_Pathurl = "tagstyle.asp?t="&t&"" '执行分页的页面
end if
Obj_Page.Setpage_Navtype = 1 '分页导航样式(可选1文字,2数字)

Call Obj_Page.Page_Init() '初始化分页

%>

<table width="99%"   align="center" cellpadding="0" cellspacing="0" class="dtable2">

        <tr class="wz">
          <td width="47">全选</td>
          <td width="47"  >编号</td>
       
          <td width="237" ><div align="center">标签标题</div></td>
          
          <td width="255"><div align="center">所属分类</div></td>
          <td width="255"><div align="center" class="STYLE2">修改操作</div></td>
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
            <input name="ID" type="checkbox" id="ID" value="<%=rs("id")%>" class="none"/>
          </div></td>
          <td ><%=rs("id")%></td>
      
        
          <td ><%=rs("tagname")%></td>
          <td ><%if t=1 then response.Write("PC端") end if%><%if t=2 then response.Write("WAP端") end if%></td>
          <td >[<a href="tagstylexg.asp?id=<%=rs("ID")%>&t=<%=t%>">修改</a>] [<a href="tagstyle_delete.asp?id=<%=rs("ID")%>&t=<%=t%>" onClick="return confirm('确定要删除此标签样式吗,前台将无法正常显示？')">删除</a>]</td>
        </tr>
  <%
rs.MoveNext
Next
	
%> 
<tr>
  <td><div align="center">
    <input type="checkbox" name="allbox" onClick="CheckAll()" class="none"/>
  </div></td>
  <td colspan="11" style="padding:5px; text-align:left">  <input class="bnt" type="submit" name="button" id="button" value="批量删除"   onClick="return confirm('确定要批量删除选中的样式吗，将不可恢复?');"/></td>  
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
	<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td >
 <div class="dmeun"><a href="biaoqian.asp" target="main">返回菜单</a></div>

 </td>
 

 </tr>
</table>
  	  <table width="99%" border="0" cellpadding="0" cellspacing="0" align="center"   class="dtable" style="margin-top:10px">
<tr>
 <td >目前暂无任何信息</td>
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