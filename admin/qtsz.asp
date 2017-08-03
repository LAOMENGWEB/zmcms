<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!-- #include file="zmwbk.asp" -->
<!-- #include file="ys.asp" -->
<%
if Instr(session("AdminPurview"),"|91,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</script>"
  response.end
end if
'========判断是否具有管理权限
%>
<%  

tjdm=request("tjdm")

killword=request("killword")
sjys=request("sjys")
zmsj=request("zmsj")

notzc=request("notzc")

If request.Form("submit") = "保存" Then



Set rs = server.CreateObject("adodb.recordset")
sql = "select * from wzset where lang='"&lang&"' "
rs.Open sql, conn, 1, 3
 rs("tjdm")= tjdm

rs("killword")=killword
rs("sjys")=sjys
rs("zmsj")=zmsj
rs("notzc")=notzc

    rs.update
    rs.Close
    Set rs = Nothing
 call addlog("其他设置")
  response.Write "<script language=javascript>alert(""恭喜您修改成功"");window.location.href=""qtsz.asp""</script>"
    response.End
End If
%>


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>zmcms_管理页面</title>
<link href="css/style.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="js/jquery-1.7.2.min.js"></script>
<style type="text/css">
table{ font-size: 14px;  border-right:none}
table td{padding:10px; font-family:"微软雅黑";color:#ccc; font-weight:bold;}
table td a{color:#ffffff}
</style>
<!--#include file="hqyy.asp"-->
</head>


<body>

<form method="POST" action="qtsz.asp" id="form" name="form">
<%
Set rs = server.CreateObject("adodb.recordset")
sql = "select * from wzset where lang='"&lang&"'"
rs.Open sql, conn, 1, 1
If rs.EOF Then
response.Write "<script language=javascript>alert(""不存在此数据！"");window.history.back();</script>"
response.End
End If
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" style=" background: #F9f9f9;border:none; ">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">其他设置  <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>


<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="set.asp" target="main" style="color:#ffffff">返回菜单</a></div></td>
 

 </tr>
</table>
<table width="99%" border="0" align="center" cellpadding="0" cellspacing="0" class="dtable4">
  
      
	<tr>
	  <td width="16%" >统计代码：</td>
      <td width="84%"> <textarea name="tjdm" cols="80" rows="6" id="tjdm" class="input_show1"><%=rs("tjdm")%></textarea></td>
    </tr>
	  <tr>
      <td >会员注册屏蔽的用户名：
        </td>
	  <td ><textarea name="notzc" cols="30" rows="5" class="input_show1"><%=rs("notzc")%></textarea>      
      (多个词之间用,隔开)</td>
    </tr>

    <tr>
      <td >留言板禁止发布脏话设置：
       </td>
	  <td > <textarea name="killword" cols="30" rows="5" class="input_show1"><%=rs("killword")%></textarea>      
      (脏话之间用,隔开)</td>
    </tr>
    

<tr> 
 <td >最近几天内信息的时间颜色：
  </td>
<td >   <input name="sjys"type="text" id="color"   value="<%=rs("sjys")%>" size="8" /> 
 <input   type="button" id="colorpicker" value="打开取色器" /> 
<input name="zmsj" type="text" id="zmsj" value="<%=rs("zmsj")%>" size="3">
天内有效
       </td>

</tr>
        
        
      <tr>
    <td >
      <input type="submit" name="Submit" value="保存" class="bnt" />
    </td>
  </tr>
</table>
</form>

</body>
</html>

<%   conn.Close
    Set conn = Nothing%>