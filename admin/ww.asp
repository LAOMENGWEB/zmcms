<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="Inc/Function.asp"-->
<!--#include file="class_page.asp"-->
<%
if Instr(session("AdminPurview"),"|91,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</script>"
  response.end
end if
'========判断是否具有管理权限
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>zmcms_管理页面</title>
<link href="css/style.css" rel="stylesheet" type="text/css">
<script language="javascript"> 
<!-- 
function CheckAll(){ 
 for (var i=0;i<eval(form1.elements.length);i++){ 
  var e=form1.elements[i]; 
  if (e.name!="allbox") e.checked=form1.allbox.checked; 
 } 
} 
--> 
</script>
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
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">旺旺客服管理  <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>

<table border="0"  cellpadding="0" cellspacing="0" width="99%" class="dtable3" align="center">
 <tr>
 <td ><div class="dmeun"><a href="kefu.asp" target="main">返回菜单</a></div>

<div class="dmeun"><a href="kfadd.asp?l=ww" target="main">旺旺客服添加</a></div>
 </td>
 

 </tr>
</table>  
<% 
If request.Form("button") = "批量删除" Then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='ww.asp';</script>" 
response.end()
end if
conn.Execute ( "delete from ww where id in ("&Request("id")&")" )
call addlog("删除旺旺客服")
response.Write "<script language=javascript>alert(""删除成功！"");window.location.href=""ww.asp""</script>"
end if

DB_Sql = "select * from ww where lang='"&lang&"' "
DB_Sql = DB_Sql & "order by px asc"
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

Obj_Page.Setpage_Pathurl = "ww.asp" '执行分页的页面
Obj_Page.Setpage_Navtype = 1 '分页导航样式(可选1文字,2数字)

Call Obj_Page.Page_Init() '初始化分页

%>

<form action="ww.asp" method="post" name="form1" id="form1"> 

            
              <table width="99%"   cellpadding="0" cellspacing="0" class="dtable2" align="center">
               
                <tr class="wz">
                  <td  >全选</td>
                  <td  >ID</td>
                  <td >排序</td>
                  <td >旺旺客服说明</td>
                  <td >旺旺号码</td>
                  <td > 操作</td>
                </tr>
              <%

'/循环开始
For II = 1 to Obj_PageListnum
If rs.Eof Then
	Exit For
End If
%>
                <tr  >
                  <td ><input name="ID" type="checkbox" id="ID" value="<%=rs("id")%>" /></td>
                  <td  ><%=rs("id")%></td>
                  <td  ><%=rs("px")%></td>
                  <td  ><%=rs("wwsm")%></td>
                  <td  ><%=rs("ww")%></td>
                  <td  >				  
                    
                    &nbsp;<a href="wwModify.asp?ID=<%=rs("ID")%>">修改</a>&nbsp; 
                    &nbsp;<a href="wwDel.asp?ID=<%=rs("ID")%>" onClick="return confirm('确定要删除该旺旺客服吗？');" >删除</a>
			</a>				  </td>
                </tr>
			
 <%
rs.MoveNext
Next
	
%>      
<tr  >
                  <td > 
                    <div >
                      <input type="checkbox" name="allbox" onClick="CheckAll()" />
      </div></td>
      <td colspan="5"  style="padding:5px; text-align:left"><input class="bnt" type="submit" name="button" id="button" value="批量删除"  onClick="return confirm('确定要批量删除选中的旺旺客服吗？');"/></td>
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


</form><%Set Obj_Page = Nothing

RS.Close
Set RS = Nothing
End If
%>
</body>
</html>
