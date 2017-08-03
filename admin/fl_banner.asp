<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<%
if Instr(session("AdminPurview"),"|111,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</script>"
  response.end
end if
'========判断是否具有管理权限
%>
<%

if request("action")="Del" then	
ID=request.QueryString("id")
set rschk=server.CreateObject("ADODB.RECORDSET")
sqlchk="select * from bannerclass where ParentClassID="+ID
rschk.open sqlchk,conn,1,1
if not rschk.eof then
response.Write "<script language=javascript>alert(""请先删除该类别的子类"");window.history.back();</script>"
response.End
end if
rschk.close
set rschk=nothing


set rsdy=server.CreateObject("ADODB.RECORDSET")
sqldy="select * from banner where zmone="&ID&" or zmtwo="&id&" "
rsdy.open sqldy,conn,1,1
if not rsdy.eof then
response.Write "<script language=javascript>alert(""请先删除该类别的下的幻灯信息"");window.history.back();</script>"
response.End
end if
rsdy.close
set rsdy=nothing

sql="Delete from bannerclass where ID="+ID
conn.execute(sql)

call addlog("删除幻灯类别")
response.Write "<script language=javascript>alert(""类别删除成功"");window.location.href=""fl_banner.asp""</script>"
	 response.End
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>zmcms_管理页面</title>
<link href="css/style.css" rel="stylesheet" type="text/css">
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
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">幻灯分类管理</div></td>
</tr>
</table>




<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="huandeng.asp" target="main">返回菜单</a></div><div class="dmeun"><a href="fladd_banner.asp" target="main">添加幻灯分类</a></div></td>
 

 </tr>
</table>


  <table width="99%" align="center" cellpadding="0" cellspacing="0"  class="dtable2">
    <tr  class="wz">
      <td >编号</td>
      <td >栏目管理</td>
      <td >排序</td>
      <td >基本操作</td>
    </tr>
    <%
			   Set rs = Server.CreateObject("ADODB.Recordset")
			   sql="select * from bannerclass where ParentClassID = 0  order by px asc"
			   rs.open sql,conn,1,1
				 
			   for i=1 to rs.recordcount
			   if rs.eof then
			   exit for
			   end if
			   str=rs("ClassName")
		 %>
		 

    <tr >
      <td width="10%" ><%=rs("ID")%></td>
      <td width="21%" ><%=str%></td>
      <td width="18%" >
          <iframe name="zmcms" width="150" height="45" frameborder="0" scrolling="no" 
				src="xgbannerflpx.asp?id=<%=rs("id")%>" ></iframe>
      </td>
      <td width="24%" ><a href="flxg_banner.asp?ID=<%=rs("ID")%>">编辑</a>/<a href="?ID=<%=rs("ID")%>&action=Del" onClick="return confirm('确定删除？')">删除</a> 
        
          <a href="banner.asp?pone=<%=rs("ID")%>"><font color="red">&nbsp;&nbsp;&nbsp;查看所属幻灯信息</font></a>
      
      </td>
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
