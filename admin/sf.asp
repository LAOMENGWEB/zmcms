<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="Inc/Function.asp"-->
<!--#include file="class_page.asp"-->
<%
if Instr(session("AdminPurview"),"|1000003,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</script>"
  response.end
end if
'========判断是否具有管理权限
%>
<% 
Result=request.QueryString("Result")
keytag=request.Form("k1")

If request.Form("button") = "设为显示" Then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='sf.asp';</script>" 
response.end()
end if

sql4="update province set xs=1 where id in ("&Request("id")&")"
conn.Execute ( sql4 )
call addlog("设置营销区域显示")
end if


If request.Form("button") = "取消显示" Then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='sf.asp';</script>" 
response.end()
end if

sql4="update province set xs=0 where id in ("&Request("id")&")"
conn.Execute ( sql4 )
call addlog("取消营销区域显示")
end if



If request.Form("del") = "批量删除" Then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='sf.asp';</script>" 
response.end()
end if
conn.Execute ( "delete from province where id in ("&Request("id")&")" )
call addlog("区域删除")
response.Write "<script language=javascript>alert(""删除成功！"");window.location.href=""sf.asp""</script>"
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
   if(confirm("确定要删除选中的区域吗？一旦删除将不能恢复！"))
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
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">区域管理  <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>


<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="chajian.asp" target="main">返回菜单</a></div><div class="dmeun"><a href="sfadd.asp" target="main">添加区域</a></div><div class="dmeun"><a href="sf.asp" target="main">区域管理</a></div></td>
 

 </tr>
</table>

<table width="99%"  cellpadding="0" cellspacing="0" class="dtable1" align="center">
  		
          <tr   >

            <form id="form2" name="form2" method="post" action="sf.asp?Result=ok">
            <td width="8%"  >
            快速搜索：</td>
            <td > <input name="k1" type="text" id="k1" />
             
              <input class="bnt" type="submit" name="button" id="button" value="搜索" /> </td>
        </form>
  </tr>
</table>

<form name="zmqy" method="Post">
 <%
if Result="ok"   then
 DB_Sql = "select * from province where province like '%"&keytag&"%' and lang='"&lang&"' "
else
DB_Sql = "select * from province where lang='"&lang&"' "
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
if Result="ok" then
Obj_Page.Setpage_Pathurl = "sf.asp?Result=ok" '执行分页的页面
else
Obj_Page.Setpage_Pathurl = "sf.asp" '执行分页的页面
end if
Obj_Page.Setpage_Navtype = 1 '分页导航样式(可选1文字,2数字)

Call Obj_Page.Page_Init() '初始化分页
%>
<table width="99%"   align="center" cellpadding="0" cellspacing="0" class="dtable2">
<tr  class="wz">
	<td width="5%"  align="center"  >全选</td>
    <td width="4%" align="center" >编号</td>
      <td width="10%" align="center" >排序</td>
    <td width="18%"  align="center" >区域名称</td>
    <td width="30%" align="center" >是否显示</td>

    <td width="14%"  align="center"  style="border-right:none">管理</td>    
  </tr>
<%
'/循环开始
For II = 1 to Obj_PageListnum
If rs.Eof Then
	Exit For
End If
%>
    <tr >
    <td  align="CENTER"><div align="center">
  <input name="ID" type="checkbox" id="ID" value="<%=rs("id")%>" class="none"/>
 </div>      </td>
    <td><div align="center"><%=rs("id")%></div></td>
         <td><div align="center"><iframe name="zmcms" width="100" height="50" frameborder="0" scrolling="no" 
				src="xgsfpx.asp?id=<%=rs("id")%>" ></iframe></div></td>
    <td ><div align="center"><%=rs("Province")%></div></td>
    <td  align="center"><%
	if rs("xs")=1 then
	response.write("显示")
	else
	response.write("不显示")
	end if
	
	%></td>

    <td align="center" style="border-right:none">
	<a href="sfxg.asp?id=<%=rs("ID")%>">编辑</a>
	<a href="delsf.asp?id=<%=rs("ID")%>" onClick="return confirm('确定要删除吗？');">删除</a>	</td>    
  </tr>

 <%
rs.MoveNext
Next	
%> 

	<tr>
    <td  align="center" ><input type="checkbox" name="allbox" onClick="CheckAll()" class="none"/></td>
    <td  colspan="6" style="padding:5px; text-align:left"><input name="Del" type="submit" class="bnt" id="Del"  onClick="JavaScript:return confirm('删除吗？')" value="批量删除">
    
    
    <input class="bnt" type="submit" name="button" id="button" value="设为显示"  onClick="return confirm('确定要显示吗？');"/>
	<input class="bnt" type="submit" name="button" id="button" value="取消显示" onClick="return confirm('确定要取消显示吗？');" />
    
    
    </td>
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
    <td >	<div align="center">
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
