<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<%
if Instr(session("AdminPurview"),"|1132,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</scrip>"
  response.end
end if
'========判断是否具有管理权限
%>
<%

if request("action")="Del" then	
ID=request.QueryString("id")
set rschk=server.CreateObject("ADODB.RECORDSET")
sqlchk="select * from yqljclass where ParentClassID="+ID
rschk.open sqlchk,conn,1,1
if not rschk.eof then
response.Write "<script language=javascript>alert(""请先删除该类别的子类"");window.history.back();</script>"
response.End
end if
rschk.close
set rschk=nothing


set rsdy=server.CreateObject("ADODB.RECORDSET")
sqldy="select * from yqlj where zmone="&ID&" "
rsdy.open sqldy,conn,1,1
if not rsdy.eof then
response.Write "<script language=javascript>alert(""请先删除该类别的下的链接信息"");window.history.back();</script>"
response.End
end if
rsdy.close
set rsdy=nothing

sql="Delete from yqljclass where ID="+ID
conn.execute(sql)

call addlog("删除链接类别")
response.Write "<script language=javascript>alert(""类别删除成功"");window.location.href=""fl_yqlj.asp""</script>"
	 response.End
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
<style type="text/css">
.table1{ font-size: 14px; background: #F9f9f9;border:none;}
.table1 td{padding:10px; font-family:"微软雅黑";color:#ccc; font-weight:bold;}
.table1 td a{color:#ffffff}
</style>
</head>

<body>



<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">链接分类管理</div></td>
</tr>
</table>



<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="lianjie.asp" target="main">返回菜单</a></div><div class="dmeun"><a href="fladd_yqlj.asp" target="main">添加链接分类</a></div></td>
 

 </tr>
</table>


  <table width="99%" align="center" cellpadding="0" cellspacing="0"  class="dtable2">
    <tr class="wz">
      <td ><div align="center" >ID</div></td>
      <td ><div align="center" >栏目管理</div></td>
      <td ><div align="center" >排序</div></td>
      <td><div align="center" >基本操作</div></td>
    </tr>
    <%
			   Set rs = Server.CreateObject("ADODB.Recordset")
			   sql="select * from yqljclass where ParentClassID = 0  order by px asc"
			   rs.open sql,conn,1,1
				 
			   for i=1 to rs.recordcount
			   if rs.eof then
			   exit for
			   end if
			   str=rs("ClassName")
		 %>
		 

    <tr >
      <td width="10%" ><div align="center"><%=rs("ID")%></div></td>
      <td width="21%" ><div align="center"><%=str%></div></td>
      <td width="18%" ><div align="center">
          <iframe name="zmcms" width="100" height="50" frameborder="0" scrolling="no" 
				src="xgyqljflpx.asp?id=<%=rs("id")%>" ></iframe>
      </div></td>
      <td width="24%" ><div align="center"><a href="flxg_yqlj.asp?ID=<%=rs("ID")%>">编辑</a>/<a href="?ID=<%=rs("ID")%>&action=Del" onClick="return confirm('确定删除？')">删除</a> 
        
          <a href="yq.asp?pone=<%=rs("ID")%>"><font color="red">&nbsp;&nbsp;&nbsp;查看所属链接信息</font></a>
      
      </div></td>
    </tr>

    
    <%
		
 
			rs.movenext
			next
			rs.close
			Set rs=nothing
    %>
   
  </table>












</body>
</html>
