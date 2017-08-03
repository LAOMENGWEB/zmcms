<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="zmwbk.asp"-->

<%
if Instr(session("AdminPurview"),"|110,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</script>"
  response.end
end if
'========判断是否具有管理权限
%>
<%
l=request.QueryString("l")
If request.Form("submit") = "确定" Then
qqsm=Trim(request.form("qqsm"))
qq=Trim(request.form("qq"))
px=Trim(request.form("px"))
Set rs = server.CreateObject("adodb.recordset")
sql = "select * from kf "
rs.Open sql, conn, 1, 3
rs.addnew
rs("qqsm") =qqsm
rs("qq") = qq
rs("px") = px
rs("lang")=lang
rs.update
rs.close
set rs=nothing
call addlog("添加QQ客服")
response.Write "<script language=javascript>alert(""恭喜您添加成功"");window.location.href=""kf.asp""</script>"
end if
If request.Form("submit") = "确定添加" Then
wwsm=Trim(request.form("wwsm"))
ww=Trim(request.form("ww"))
px=Trim(request.form("px"))
Set rs = server.CreateObject("adodb.recordset")
sql = "select * from ww "
rs.Open sql, conn, 1, 3
rs.addnew
rs("wwsm") =wwsm
rs("ww") = ww
rs("px") = px
rs("lang")=lang
rs.update
rs.close
set rs=nothing
call addlog("添加旺旺客服")
response.Write "<script language=javascript>alert(""恭喜您添加成功"");window.location.href=""ww.asp""</script>"
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
<%if l="qq" then%>



<form   method="post" name="myform" >
<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">QQ客服添加  <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>

<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="kefu.asp" target="main" style="color:#ffffff">返回菜单</a></div>

<div class="dmeun"><a href="kf.asp" target="main" style="color:#ffffff">QQ客服管理</a></div>

 </td>



 </tr>
</table>
                <table width="99%"  border="0" cellpadding="0" cellspacing="0"  align="center" class="dtable4">
                  
				  <tr  > 
                    <td width="11%"> QQ客服说明：</td>
					<td width="89%">
					<input name="qqsm" type="text" class="inputtext" id="qqsm" size="25" maxlength="30" />
                    </td>
                  </tr>
				  <tr  > 
                    <td> QQ号码：</td>
					<td><input name="qq"type="text" class="inputtext" id="qq"   size="20" maxlength="30"></td>
                  </tr>
				  
				  
                  
                  
                  <tr  >
                    <td >排序：</td>
					<td>
                      <input name="px" type="text" id="px"  style="width:30px" value="1"/>
                    越小越靠前</td>
                  </tr>
                  
                  
                  
                  
                  
                  <tr  > 
                    <td colspan="2"> 
                        <input class="bnt"name="submit" type="submit" value="确定">
                                         </td>
                  </tr>
            </table>
            
          </form>
       

  <%end if%>
  <%if l="ww" then%>

<form   method="post" name="myform" >
<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">旺旺客服添加  <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>

<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="kefu.asp" target="main" style="color:#ffffff">返回菜单</a></div>

<div class="dmeun"><a href="ww.asp" target="main" style="color:#ffffff">旺旺客服管理</a></div>

 </td>
 

 </tr>
</table>

                <table width="99%"  border="0" cellpadding="0" cellspacing="0"  align="center" class="dtable4">
                  
				  <tr  > 
                    <td width="9%"> 旺旺说明：</td>
					<td width="91%"><input name="wwsm" type="text" class="inputtext" id="wwsm" size="25" maxlength="30" />
                    </td>
                  </tr>
				  <tr  > 
                    <td> 旺旺号码：</td>
					<td><input name="ww"type="text" class="inputtext" id="ww"   size="20" maxlength="30"></td>
                  </tr>
				  
				  
                  
                  
                  <tr  >
                    <td>排序：</td>
					<td>
                      <input name="px" type="text" id="px" value="0" style="width:30px"/>
                    越小越靠前</td>
                  </tr>
                  
                  
                  
                  
                  
                  <tr  > 
                    <td colspan="2"> 
                        <input class="bnt"name="submit" type="submit" value="确定添加">
                    </td>
                  </tr>
            </table>
            
          </form>
       <%end if%>

</body>
</html>
