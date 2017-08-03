<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!-- #include file="Inc/function.asp" -->
<!--#include file="sc.asp"-->
<%
if Instr(session("AdminPurview"),"|91,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</script>"
  response.end
end if
'========判断是否具有管理权限
%>
<%  


mde=request("mde")
minfo=request("minfo")
mpro=request("mpro")
mph=request("mph")
mtv=request("mtv")
mdo=request("mdo")
mzp=request("mzp")
m1=request("m1")
m2=request("m2")
m3=request("m3")
m4=request("m4")
m6=request("m6")
If request.Form("submit") = "保存" Then
Set rs = server.CreateObject("adodb.recordset")
sql = "select * from wzset where lang='"&lang&"' "
rs.Open sql, conn, 1, 3
rs("mde")=mde
rs("minfo")=minfo
rs("mpro")=mpro
rs("mph")=mph
rs("mtv")=mtv
rs("mdo")=mdo
rs("mzp")=mzp
rs("m1")=m1
rs("m2")=m2
rs("m3")=m3
rs("m4")=m4
rs("m6")=m6

    rs.update
    rs.Close
    Set rs = Nothing
   call addlog("修改菜单高亮值")
  response.Write "<script language=javascript>alert(""恭喜您修改成功"");window.location.href=""hight.asp""</script>"
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
.table1{ font-size: 14px; background: #F9f9f9;border:none;}
.table1 td{padding:10px; font-family:"微软雅黑";color:#ccc; font-weight:bold;}
.table1 td a{color:#ffffff}
</style>
<!--#include file="hqyy.asp"-->
</head>


<body>

<form method="POST" action="hight.asp" id="form" name="form">
<%
Set rs = server.CreateObject("adodb.recordset")
sql = "select * from wzset where lang='"&lang&"'"
rs.Open sql, conn, 1, 1

%>
<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">菜单高亮值设置  <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>

<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="set.asp" target="main" STYLE="color:#ffffff">返回菜单</a></div></td>
 

 </tr>
</table>
<table width="99%" border="0" align="center" cellpadding="0" cellspacing="0" class="dtable4">
  
        <tr>
          <td width="15%" >默认首页高亮值：</td>
          <td width="85%" ><input name="mde" type="text" id="mde" value="<%=rs("mde")%>" size="3"></td>
        </tr>
        <tr>
          <td >新闻相关高亮值：
		    </td>
          <td ><input name="minfo" type="text" id="minfo" value="<%=rs("minfo")%>" size="3"></td>
        </tr>
        <tr>
          <td >产品相关高亮值：
		    </td>
          <td ><input name="mpro" type="text" id="mpro" value="<%=rs("mpro")%>" size="3"></td>
        </tr>
        <tr>
          <td >图片相关高亮值：
            </td>
          <td ><input name="mph" type="text" id="mph" value="<%=rs("mph")%>" size="3"></td>
        </tr>
        <tr>
          <td >下载相关高亮值：
            </td>
          <td ><input name="mdo" type="text" id="mdo" value="<%=rs("mdo")%>" size="3"></td>
        </tr>
        <tr>
          <td >视频相关高亮值：
            </td>
          <td ><input name="mtv" type="text" id="mtv" value="<%=rs("mtv")%>" size="3"></td>
        </tr>
        <tr>
          <td >招聘相关高亮值：
          </td>
          <td ><input name="mzp" type="text" id="mzp" value="<%=rs("mzp")%>" size="3"></td>
        </tr>
        
        
        
         <tr>
          <td >购物车高亮值：
          </td>
          <td ><input name="m1" type="text" id="m1" value="<%=rs("m1")%>" size="3"></td>
        </tr>
        
         <tr>
          <td >会员中心高亮值  针对以后支付按钮：
          </td>
          <td ><input name="m2" type="text" id="m2" value="<%=rs("m2")%>" size="3"></td>
        </tr>
        
         <tr>
          <td >银行转账页面高亮值：
          </td>
          <td ><input name="m3" type="text" id="m3" value="<%=rs("m3")%>" size="3"></td>
        </tr>
        
         <tr>
          <td >qq跳转页面高亮值：
          </td>
          <td ><input name="m6" type="text" id="m6" value="<%=rs("m6")%>" size="3"></td>
        </tr>
        
             <tr>
          <td >支付成功后返回的m值：
          </td>
          <td ><input name="m4" type="text" id="m4" value="<%=rs("m4")%>" size="3"></td>
        </tr>
        
        
        
        
        
        
        
        
		
    <td colspan="2" >
      <input type="submit" name="Submit" value="保存" class="bnt" />  </td>
  </tr>
</table>
</form>

</body>
</html>

<% conn.Close
    Set conn = Nothing%>