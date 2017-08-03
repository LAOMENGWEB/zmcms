<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="Inc/Function.asp"-->
<%
if Instr(session("AdminPurview"),"|101,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</script>"
response.end
end if
'========判断是否具有管理权限
t=request.QueryString("t")
If request.Form("button") = "设为显示" Then
if Request("chk")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');;window.history.back();</script>" 
response.end()
end if
dim sql4 
sql4="update meun set xs=1 where id in ("&Request("chk")&")"
conn.Execute ( sql4 )
call addlog("导航管理——设为显示")
end if
If request.Form("button") = "取消显示" Then
if Request("chk")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.history.back();</script>" 
response.end()
end if
dim sql2 
sql2="update meun set xs=2 where id in ("&Request("chk")&")"
conn.Execute ( sql2 )
call addlog("导航管理——取消显示")
end if
ID=checknum(request.QueryString("ID"),0)
AllID=request("chk")
if request("action")="Del" then
set rschk=server.CreateObject("ADODB.RECORDSET")
	sqlchk="select * from meun where ParentClassID="+ID
	rschk.open sqlchk,conn,1,1
	if not rschk.eof then
	response.Write "<script language=javascript>alert(""请先删除该类别的子类"");window.history.back();</script>"
    response.End
	end if
	rschk.close
	set rschk=nothing
	set rschk1=server.CreateObject("ADODB.RECORDSET")
	sqlchk1="select * from meun where ID="+ID
	rschk1.open sqlchk1,conn,1,1
	ParentClassID1=checknum(rschk1("ParentClassID"),0)
	SmallImage=rschk1("SmallImage")
	rschk1.close
	set rschk1=nothing
	call deletefile1(SmallImage)
	sql="Delete from meun where ID="+ID
	conn.execute(sql)
    ChildCount=conn.execute("select count(*) from [meun] where ParentClassID="&ParentClassID1)(0)
    conn.execute("update [meun] set c_child = "& checknum(ChildCount,0) &" where id="& ParentClassID1 &"")
	
	if t=1 Then
		call addlog("删除顶部导航菜单")
	response.Write "<script language=javascript>alert(""类别删除成功"");window.location.href=""menugl.asp?t=1""</script>"
	 response.End
	 end if
	 if t=0 Then
		 call addlog("删除底部导航菜单")
	 response.Write "<script language=javascript>alert(""类别删除成功"");window.location.href=""menugl.asp?t=0""</script>"
	 response.End
	 end if
	  if t=2 Then
		  call addlog("删除其他导航菜单")
	 response.Write "<script language=javascript>alert(""类别删除成功"");window.location.href=""menugl.asp?t=2""</script>"
	 response.End
	 end If
	 	  if t=3 Then
		 	  call addlog("删除手机导航菜单")
	 response.Write "<script language=javascript>alert(""类别删除成功"");window.location.href=""menugl.asp?t=3""</script>"
	 response.End
	 end if
end if
if request("action")="delall" then
	If AllID="" Then response.write"<script language=javascript>alert('没有选择数据');history.go(-1);</script>":response.End
	arrAllID=split(AllID,",")
	for arrAllID_i=0 to ubound(arrAllID)
	set rschk1=server.CreateObject("ADODB.RECORDSET")
	sqlchk1="select * from meun where ID="+arrAllID(arrAllID_i)
	rschk1.open sqlchk1,conn,1,1
	ParentClassID1=checknum(rschk1("ParentClassID"),0)
	SmallImage=rschk1("SmallImage")
	rschk1.close
	set rschk1=nothing
	call deletefile1(SmallImage)
	sql="delete from meun where ID ="&arrAllID(arrAllID_i)&""
	conn.execute(sql)
	ChildCount=conn.execute("select count(*) from [meun] where ParentClassID="&ParentClassID1)(0)
    conn.execute("update [meun] set c_child = "& checknum(ChildCount,0) &" where id="& ParentClassID1 &"")
	next
	if t=1 then
	call addlog("批量删除顶部菜单")
	response.Write "<script language=javascript>alert(""删除成功"");window.location.href=""menugl.asp?t=1""</script>"
	end if
	if t=0 then
	call addlog("批量删除底部菜单")
	response.Write "<script language=javascript>alert(""删除成功"");window.location.href=""menugl.asp?t=0""</script>"
	end if
	if t=2 then
	call addlog("批量删除其他菜单")
	response.Write "<script language=javascript>alert(""删除成功"");window.location.href=""menugl.asp?t=2""</script>"
	end If
		if t=3 then
	call addlog("批量删除WAP菜单")
	response.Write "<script language=javascript>alert(""删除成功"");window.location.href=""menugl.asp?t=3""</script>"
	end if
	response.end
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>zmcms_管理页面</title>
<link href="css/style.css" rel="stylesheet" type="text/css">
		
		
<script language="javascript" type="text/javascript">
function selcheck(id) {
   var e =document.getElementById(id);
   var objs = e.getElementsByTagName("input");
   if (e.checked){
      e.checked = false;
      for(var i=0; i<objs.length; i++) {
         if(objs[i].type.toLowerCase() == "checkbox" )
         objs[i].checked = false;
      }
   }   else   {
      e.checked = true;
      for(var i=0; i<objs.length; i++) {
         if(objs[i].type.toLowerCase() == "checkbox" )
         objs[i].checked = true;
      }
   }
}
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
<form method="POST" id=form name="form">
<div id="content">


<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center"> 
 <%if t=1 then%>顶部导航管理<%end if%>
 <%if t=0 then%>底部导航管理<%end if%>
 <%if t=2 then%>其他导航管理<%end if%>
 <%if t=3 then%>手机导航管理<%end if%>  <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>




<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="daohang.asp" target="main">返回菜单</a></div>
 <div class="dmeun">
 <%if t=1 then%>
 <a href="menuadd.asp?t=1" target="main">添加顶部导航</a>
 <%end if%>
 <%if t=0 then%>
 <a href="menuadd.asp?t=0" target="main">添加底部导航</a>
 <%end if%>
 <%if t=2 then%>
 <a href="menuadd.asp?t=2" target="main">添加其他导航</a>
 <%end if%>
 <%if t=3 then%>
 <a href="menuadd.asp?t=3" target="main">添加手机导航</a>
 <%end if%>
 </div>
 </td>
 

 </tr>
</table>
  <table width="99%" align="center" cellpadding="0" cellspacing="0"  class="dtable2">
    <tr  class="wz">
      <td   ><div align="center" class="STYLE2">全选</div></td>
      <td ><div align="center" class="STYLE2">ID</div></td>
      <td ><div align="center" class="STYLE2">菜单管理</div></td>

      <td ><div align="center" class="STYLE2">是否显示</div></td>
      <td ><div align="center" class="STYLE2">排序</div></td>
      <td style="border-right:none"><div align="center" class="STYLE2" >基本操作</div></td>
    </tr>
    <%
			   Set rs = Server.CreateObject("ADODB.Recordset")
			   sql="select * from meun where ParentClassID = 0 and fl="&t&" and lang='"&lang&"' order by px asc"
			   rs.open sql,conn,1,1
				 
			   for i=1 to rs.recordcount
			   if rs.eof then
			   exit for
			   end if
			   str=rs("ClassName")
		 %>
    <tr >
      <td width="5%" ><div align="center">
          <input type="checkbox" name="chk" value="<%=rs("ID")%>" style="border:none"/>
      </div></td>
      <td width="5%" ><div align="center"><%=rs("ID")%></div></td>
      <td width="24%" ><div align="left" style="padding-left:10px">|-&nbsp;&nbsp;<font color="<%=rs("mcolor")%>"><%=str%></font>
          <%if rs("bieming")<>"" then%>
          (<%=rs("bieming")%>)
          <%end if%>
      </div></td>
  
      <td width="12%" ><div align="center">
          <% if rs("xs")=1 then  Response.Write "<font color=red>首页显示</font>"  else Response.Write "<font color=red>不显示</font>"  end if%>
      </div></td>
      <td width="14%" ><div align="center">
          <iframe name="zmcms" width="150" height="50" frameborder="0" scrolling="no" 
				src="xgdhpx.asp?id=<%=rs("id")%>" ></iframe>
      </div></td>
      <td width="14%" style="border-right:none"><div align="center">
	  <%if t=1 then%>
	  <a href="menuadd.asp?ID=<%=rs("ID")%>&t=1">添加子栏目</a>
	  <%end if%>
	  <%if t=0 then%>
	  <a href="menuadd.asp?ID=<%=rs("ID")%>&t=0">添加子栏目</a>
	  <%end if%>
	  <%if t=2 then%>
	  <a href="menuadd.asp?ID=<%=rs("ID")%>&t=2">添加子栏目</a>
	  <%end if%>
	   <%if t=3 then%>
	  <a href="menuadd.asp?ID=<%=rs("ID")%>&t=3">添加子栏目</a>
	  <%end if%>
	  丨
        <%if t=1 then%>
        <a href="flxg_menu.asp?ID=<%=rs("ID")%>&t=1">编辑</a>
        <%end if%>
		<%if t=0 then%>
        <a href="flxg_menu.asp?ID=<%=rs("ID")%>&t=0">编辑</a>
              <%end if%>
			<%if t=2 then%>
        <a href="flxg_menu.asp?ID=<%=rs("ID")%>&t=2">编辑</a>
              <%end if%>  
			  	<%if t=3 then%>
        <a href="flxg_menu.asp?ID=<%=rs("ID")%>&t=3">编辑</a>
              <%end if%>  
			  
        丨
     <%if t=1 then%>
		<a href="?ID=<%=rs("ID")%>&action=Del&t=1" onClick="return confirm('确定删除？')">删除</a>
	<%end if%>
	
	  <%if t=0 then%>
		<a href="?ID=<%=rs("ID")%>&action=Del&t=0" onClick="return confirm('确定删除？')">删除</a>
	<%end if%>
	  <%if t=2 then%>
		<a href="?ID=<%=rs("ID")%>&action=Del&t=2" onClick="return confirm('确定删除？')">删除</a>
	<%end if%>
	  <%if t=3 then%>
		<a href="?ID=<%=rs("ID")%>&action=Del&t=3" onClick="return confirm('确定删除？')">删除</a>
	<%end if%>
		
		</div></td>
    </tr>
    <%
			   Set rs1 = Server.CreateObject("ADODB.Recordset")
			   sql1="select * from meun where ParentClassID = "& rs("ID") &" and fl="&t&" and lang='"&lang&"' order by px asc"
			   rs1.open sql1,conn,1,1
			   
			   for j=1 to rs1.recordcount
			   if rs1.eof then
			   exit for
			   end if
			   str1=rs1("ClassName")
		 %>
    <tr >
      <td  ><div align="center">
          <input type="checkbox" name="chk" value="<%=rs1("ID")%>" />
      </div></td>
      <td  ><div align="center"><%=rs1("ID")%></div></td>
      <td  ><div align="left" style="padding-left:30px">|-|-  <font color="<%=rs1("mcolor")%>"><%=str1%><%if rs1("bieming")<>"" then%>(<%=rs1("bieming")%>)<%end if%></font></div></td>

      <td  ><div align="center">
          <% if rs1("xs")=1 then  Response.Write "<font color=red>首页显示"  else Response.Write "<font color=red>不显示</font>"  end if%>
      </div></td>
      <td  ><div align="center">
        <div align="center">
            <iframe name="zmcms" width="150" height="50" frameborder="0" scrolling="no" 
				src="xgdhpx.asp?id=<%=rs1("id")%>" ></iframe>
        </div>
      </div></td>
      <td  style="border-right:none"><div align="center">
	  
	    <%if t=1 then%>
	  <a href="menuadd.asp?ID=<%=rs1("ID")%>&t=1">添加子栏目</a>
	  <%end if%>
	  <%if t=0 then%>
	  <a href="menuadd.asp?ID=<%=rs1("ID")%>&t=0">添加子栏目</a>
	  <%end if%>
	  <%if t=2 then%>
	  <a href="menuadd.asp?ID=<%=rs1("ID")%>&t=2">添加子栏目</a>
	  <%end if%>
	   <%if t=3 then%>
	  <a href="menuadd.asp?ID=<%=rs1("ID")%>&t=3">添加子栏目</a>
	  <%end if%>
	  丨
	  
	  
	  
          <%if t=1 then%>
          <a href="flxg_menu.asp?ID=<%=rs1("ID")%>&t=1">编辑</a>
          <%end if%>
          <%if t=0 then%>
          <a href="flxg_menu.asp?ID=<%=rs1("ID")%>&t=0">编辑</a>
          <%end if%>
		  <%if t=2 then%>
          <a href="flxg_menu.asp?ID=<%=rs1("ID")%>&t=2">编辑</a>
          <%end if%>
		  <%if t=3 then%>
          <a href="flxg_menu.asp?ID=<%=rs1("ID")%>&t=3">编辑</a>
          <%end if%>
        丨
		<%if t=1 then%>
		<a href="?ID=<%=rs1("ID")%>&action=Del&t=1" onClick="return confirm('确定删除？')">删除</a>
	<%end if%>
	
	  <%if t=0 then%>
		<a href="?ID=<%=rs1("ID")%>&action=Del&t=0" onClick="return confirm('确定删除？')">删除</a>
	<%end if%>
	  <%if t=2 then%>
		<a href="?ID=<%=rs1("ID")%>&action=Del&t=2" onClick="return confirm('确定删除？')">删除</a>
	<%end if%>
		 <%if t=3 then%>
		<a href="?ID=<%=rs1("ID")%>&action=Del&t=3" onClick="return confirm('确定删除？')">删除</a>
	<%end if%>
		
		</div></td>
    </tr>
	
	
	
	
	
	   <%
			   Set rs2 = Server.CreateObject("ADODB.Recordset")
			   sql2="select * from meun where ParentClassID = "& rs1("ID") &" and fl="&t&" and lang='"&lang&"' order by px asc"
			   rs2.open sql2,conn,1,1
			   
			   for k=1 to rs2.recordcount
			   if rs2.eof then
			   exit for
			   end if
			   str2=rs2("ClassName")
		 %>
    <tr >
      <td  ><div align="center">
          <input type="checkbox" name="chk" value="<%=rs2("ID")%>" />
      </div></td>
      <td  ><div align="center"><%=rs2("ID")%></div></td>
      <td  ><div align="left" style="padding-left:55px">|-|-|-  <font color="<%=rs2("mcolor")%>"><%=str2%></font><%if rs2("bieming")<>"" then%>(<%=rs2("bieming")%>)<%end if%></div></td>

      <td  ><div align="center">
          <% if rs2("xs")=1 then  Response.Write "<font color=red>首页显示"  else Response.Write "<font color=red>不显示</font>"  end if%>
      </div></td>
      <td  ><div align="center">
        <div align="center">
            <iframe name="zmcms" width="150" height="50" frameborder="0" scrolling="no" 
				src="xgdhpx.asp?id=<%=rs2("id")%>" ></iframe>
        </div>
      </div></td>
      <td  style="border-right:none"><div align="center">
          <%if t=1 then%>
          <a href="flxg_menu.asp?ID=<%=rs2("ID")%>&t=1">编辑</a>
          <%end if%>
          <%if t=0 then%>
          <a href="flxg_menu.asp?ID=<%=rs2("ID")%>&t=0">编辑</a>
          <%end if%>
		  <%if t=2 then%>
          <a href="flxg_menu.asp?ID=<%=rs2("ID")%>&t=2">编辑</a>
          <%end if%>
		  <%if t=3 then%>
          <a href="flxg_menu.asp?ID=<%=rs2("ID")%>&t=3">编辑</a>
          <%end if%>
        丨
		<%if t=1 then%>
		<a href="?ID=<%=rs2("ID")%>&action=Del&t=1" onClick="return confirm('确定删除？')">删除</a>
	<%end if%>
	
	  <%if t=0 then%>
		<a href="?ID=<%=rs2("ID")%>&action=Del&t=0" onClick="return confirm('确定删除？')">删除</a>
	<%end if%>
	  <%if t=2 then%>
		<a href="?ID=<%=rs2("ID")%>&action=Del&t=2" onClick="return confirm('确定删除？')">删除</a>
	<%end if%>
		  <%if t=3 then%>
		<a href="?ID=<%=rs2("ID")%>&action=Del&t=3" onClick="return confirm('确定删除？')">删除</a>
	<%end if%>
		
		</div></td>
    </tr>
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
    <%
		rs2.movenext
			next
			rs2.close
			set rs2=nothing
	
	
			rs1.movenext
			next
			rs1.close
			set rs1=nothing
 
			rs.movenext
			next
			rs.close
			Set rs=nothing
    %>
    <tr >
      <td style="border-bottom:none"><div align="center">
          <input type="checkbox" name="chkall" onClick="selcheck('content')" />
      <a href="javascript:selcheck('content');"></a></div></td>
      <td colspan="6"  style="padding:5px; text-align:left">
	  <%if t=1 then%>
          <input class="bnt" type="submit" name="Delete_Button" value="批量删除" onClick="document.form.action='?action=delall&t=1';return confirm('确定要删除吗?');" />
		  <%end if%>
		   <%if t=2 then%>
          <input class="bnt" type="submit" name="Delete_Button" value="批量删除" onClick="document.form.action='?action=delall&t=2';return confirm('确定要删除吗?');" />
		  <%end if%>
		   <%if t=0 then%>
          <input class="bnt" type="submit" name="Delete_Button" value="批量删除" onClick="document.form.action='?action=delall&t=0';return confirm('确定要删除吗?');" />
		  <%end if%>
		    <%if t=3 then%>
          <input class="bnt" type="submit" name="Delete_Button" value="批量删除" onClick="document.form.action='?action=delall&t=3';return confirm('确定要删除吗?');" />
		  <%end if%>
      
        <input class="bnt" type="submit" name="button" id="button" value="设为显示" onClick="return confirm('确定要显示选中的导航栏目吗？');"/>
        <input class="bnt" type="submit" name="button" id="button" value="取消显示" onClick="return confirm('确定要取消显示选中的导航栏目吗？');"/></td>
    </tr>
  </table>
</div>
</form>
</body>
</html>
