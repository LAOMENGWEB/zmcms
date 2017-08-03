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
mode=request.QueryString("mode")
typeid=request.QueryString("typeid")
If request.Form("button") = "批量删除" Then
if Request("sid")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='selectglvaule.asp?id="&id&"&mode="&mode&"';</script>" 
response.end()
end if

conn.Execute ( "delete from SelectValue where id in ("&Request("sid")&")" )
call addlog("批量删除属性值")
end if


%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>属性值管理</title>

        
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
   if(confirm("确定要删除选中的属性值吗？一旦删除将不能恢复！"))
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



   
<form name="zmqy" method="Post">






<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">   <%if mode=0 then%>筛选属性值管理<%end if%>
    <%if mode=1 then%>规格属性值管理<%end if%>
    <%if mode=2 then%>商品属性值管理<%end if%>  <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>


<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="shuxing.asp" target="main">返回菜单</a></div>

<%if mode=0 then%> 
<div class="dmeun"><a href="select.asp?mode=0" target="main">筛选属性管理</a></div>
<%end if%>
<%if mode=1 then%> 
<div class="dmeun"><a href="select.asp?mode=1" target="main">规格属性管理</a></div>
<%end if%>
<%if mode=2 then%> 
<div class="dmeun"><a href="select.asp?mode=2" target="main">商品属性管理</a></div>
<%end if%>

<%if mode=0 then%> 
<div class="dmeun"><a href="selectaddvalue.asp?s_id=<%=id%>&mode=<%=mode%>&typeid=<%=typeid%>" target="main">添加筛选属性值</a></div>
<%end if%>
<%if mode=1 then%> 
<div class="dmeun"><a href="selectaddvalue.asp?s_id=<%=id%>&mode=<%=mode%>&typeid=<%=typeid%>" target="main">添加规格属性值</a></div>
<%end if%>
<%if mode=2 then%> 
<div class="dmeun"><a href="selectaddvalue.asp?s_id=<%=id%>&mode=<%=mode%>&typeid=<%=typeid%>" target="main">添加商品属性值</a></div>
<%end if%>
 </td>
 

 </tr>
</table>


<%

DB_Sql = "select * from SelectValue where s_id="&id&" and s_mode="&mode&" and lang='"&lang&"' "


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



Obj_Page.Setpage_Pathurl = "selectglvaule.asp?id="&id&" and mode="&mode&"" '执行分页的页面


Obj_Page.Setpage_Navtype = 1 '分页导航样式(可选1文字,2数字)

Call Obj_Page.Page_Init() '初始化分页
%>

  <table width="99%" align="center" cellpadding="0" cellspacing="0" class="dtable2" >
  		
          <tr class="wz">
            <td  ><div align="center">全选</div></td>
            <td   ><div align="center">编号</div></td>
            <td   ><div align="center">所属属性</div></td>
            <td   ><div align="center">属性值名称</div></td>
            <td  ><div align="center">管理操作</div></td>
          </tr>
<%
'/循环开始
For II = 1 to Obj_PageListnum
If rs.Eof Then
	Exit For
End If
%>
          <tr   > 
            <td > 
              <div align="center">
  <input name="sID" type="checkbox" id="sID" value="<%=rs("id")%>" style="border:none"/>
 </div></td>
		    <td ><div align="center"><%=rs("id")%></div></td>
		    <td ><div align="center">
		      <%dim rsc
		  set rsc=server.CreateObject("adodb.recordset")
		  rsc.open "select * from SelectName where lang='"&lang&"'",conn,1,1
		  while not rsc.eof
		    if rs("s_id")=rsc("s_id") then
			response.Write("" & rsc("s_name") & "")
			end if
			
			rsc.movenext
		wend
		rsc.close
		set rsc=nothing
		  %>
	        </div></td>
		    <td ><div align="center"><%= (rs("s_value")) %></div></td>
		    <td ><div align="center"><a href="selectxgvaule.asp?ID=<%=rs("id")%>&mode=<%=mode%>&s_id=<%=id%>">修改</a> 
		          <a href="selectdelvaule.asp?ID=<%=rs("ID")%>&mode=<%=mode%>&s_id=<%=id%>" onClick="return ConfirmDel();">删除  </a>
	            </div></td>
          </tr>
		      <%
rs.MoveNext
Next	
%>          <tr    >
           <td ><div align="center">
             <input type="checkbox" name="allbox" onClick="CheckAll()" style="border:none"/>
           </div></td>
           <td colspan="5" style="padding:5px; text-align:left">
             <input class="bnt" type="submit" name="button" id="button" value="批量删除"   onClick="return confirm('确定要批量删除选中的属性值吗，将不可恢复?');"/>
&nbsp;</td>
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
	<table width="99%" align="center" cellpadding="0" cellspacing="0" class="table1" style="margin-top:10px">
  		
          <tr   >
            <td>  <%Response.Write "目前暂无任何信息"
Response.end%> </td>       
  </table>

	
<%Set Obj_Page = Nothing

RS.Close
Set RS = Nothing

End If

%>
 
</form>














</body>
</html>
