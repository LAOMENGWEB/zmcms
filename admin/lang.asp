<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="Inc/Function.asp"-->
<!--#include file="class_page.asp"-->
<%
if Instr(session("AdminPurview"),"|1000001,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</script>"
  response.end
end if
'========判断是否具有管理权限
If Request("x")=1 Then

sql4="update lang set xs=1 where id = "&Request("id")&""
conn.Execute ( sql4 )
End If
If Request("x")=0 Then

sql4="update lang set xs=0 where id = "&Request("id")&""
conn.Execute ( sql4 )
End if

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
   if(confirm("确定要删除选中的语言吗？一旦删除将不能恢复！"))
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

<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">多语言管理</div></td>
</tr>
</table>





<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="langadd.asp" target="main">添加语言</a></div></td>
 

 </tr>
</table>



<form name="zmqy" method="Post">
 <%
DB_Sql = "select * from lang  "
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

Obj_Page.Setpage_Pathurl = "lang.asp" '执行分页的页面
Obj_Page.Setpage_Navtype = 1 '分页导航样式(可选1文字,2数字)

Call Obj_Page.Page_Init() '初始化分页
%>
<table width="99%"   align="center" cellpadding="0" cellspacing="0" class="dtable2">
<tr  class="wz">
	<td width="5%"  align="center"  >全选</td>
    <td width="4%" align="center" >编号</td>
    <td width="10%" align="center" >排序</td>
    <td width="10%"  align="center" >语言名称</td>
    <td width="10%" align="center" >语言标识</td>
     <td width="5%" align="center" >语言图标</td>
    
    <td width="5%" align="center" >货币符号</td>
    <td width="5%" align="center" >货币文字</td>
     <td width="5%" align="center" >是否显示</td>
     <td width="7%"  align="center" >前台默认语言</td>
    <td width="7%"  align="center" >后台默认语言</td>
    <td width="30%"  align="center"  style="border-right:none">管理</td>    
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
				src="xglangpx.asp?id=<%=rs("id")%>" ></iframe></div></td>
    <td ><div align="center"><%=rs("langname")%></div></td>
    <td  align="center"><%=rs("langbs")%></td>
     <td  align="center">
	     <%
	     If rs("SmallImage")<>"" Then 
	     Response.Write "<img src='"&sitepath&""&rs("SmallImage")&"'>"
         Else
	         Response.Write("无图标")
         End if


	     %></td>
    <td  align="center"><%=rs("huobi")%></td>
    <td  align="center"><%=rs("huobi1")%></td>
     <td  align="center"><%
    If rs("xs")=1 Then 
    Response.Write ("是")
	Else
	Response.Write ("否")
    End if
	    %></td>
    
    <td  align="center"><%
    If rs("qt")=1 Then 
    Response.Write ("是")
	Else
	Response.Write ("否")
    End if
	    %></td>
    <td  align="center"><%
    If rs("ht")=1 Then 
    Response.Write ("是")
	Else
	Response.Write ("否")
    End if
	    %></td>
   
    <td align="center" style="border-right:none">
	<a href="langmr.asp?id=<%=rs("ID")%>&l=q" onClick="return confirm('确定要设为前台默认语言？');">设为前台默认语言</a>
	<a href="langmr.asp?id=<%=rs("ID")%>&l=h" onClick="return confirm('确定要设为后台默认语言？');">设为后台默认语言</a>
	<%If rs("xs")=0 then%>
	<a href="?id=<%=rs("ID")%>&x=1" onClick="return confirm('确定要设为显示？');">设为显示</a>
	<%End if%>
	<%If rs("xs")=1 then%>
	<a href="?id=<%=rs("ID")%>&x=0" onClick="return confirm('确定要取消显示？');">取消显示</a>
	<%End if%>
	<a href="langxg.asp?id=<%=rs("ID")%>">编辑</a>
	<a href="dellang.asp?id=<%=rs("ID")%>" onClick="return confirm('确定要删除吗？');">删除</a>	</td>    
  </tr>

 <%
rs.MoveNext
Next	
%> 


</table>
<%If Obj_Page.OutErrInfo <> "" Then
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

<%Set Obj_Page = Nothing

RS.Close
Set RS = Nothing
End If
%>
</form>
</body>
</html>
