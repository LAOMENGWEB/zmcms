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

newinfo=request("newinfo")
ProductInfo=request("ProductInfo")
alinfo=request("alinfo")
downinfo=request("downinfo")
tvinfo=request("tvinfo")
zpinfo=request("zpinfo")
lyinfo=request("lyinfo")
hdsl=request("hdsl")
hylyinfo=request("hylyinfo")
sodd=request("sodd")
hyypinfo=request("hyypinfo")
hyorderinfo=request("hyorderinfo")
yxwlinfo=request("yxwlinfo")
If request.Form("submit") = "保存" Then
Set rs = server.CreateObject("adodb.recordset")
sql = "select * from wzset where lang='"&lang&"' "
rs.Open sql, conn, 1, 3
rs("newinfo")=newinfo
rs("ProductInfo")=ProductInfo
rs("alinfo")=alinfo
rs("downinfo")=downinfo
rs("tvinfo")=tvinfo
rs("zpinfo")=zpinfo
rs("sodd")=sodd
rs("lyinfo")=lyinfo
rs("hdsl")=hdsl
rs("hylyinfo")=hylyinfo
rs("hyypinfo")=hyypinfo
rs("hyorderinfo")=hyorderinfo
rs("yxwlinfo")=hyorderinfo
rs("lang")=lang
    rs.update
    rs.Close
    Set rs = Nothing
   call addlog("修改列表页显示数量")
  response.Write "<script language=javascript>alert(""恭喜您修改成功"");window.location.href=""lbsm.asp""</script>"
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

<form method="POST" action="lbsm.asp" id="form" name="form">
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
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">列表页数量设置  <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>
<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="set.asp" target="main" STYLE="color:#ffffff">返回菜单</a></div></td>
 

 </tr>
</table>
<table width="99%" border="0" align="center" cellpadding="0" cellspacing="0" class="dtable4">
  
        <tr>
          <td width="15%" >幻灯片显示数量：</td>
          <td width="85%" ><input name="hdsl" type="text" id="hdsl" value="<%=rs("hdsl")%>" size="3"></td>
        </tr>
        <tr>
          <td >新闻列表页调用条数：
		    </td>
          <td ><input name="newinfo" type="text" id="newinfo" value="<%=rs("newinfo")%>" size="3"></td>
        </tr>
        <tr>
          <td >产品列表页调用条数：
		    </td>
          <td ><input name="ProductInfo" type="text" id="ProductInfo" value="<%=rs("productinfo")%>" size="3"></td>
        </tr>
        <tr>
          <td >图片列表页调用条数：
            </td>
          <td ><input name="alinfo" type="text" id="alinfo" value="<%=rs("alinfo")%>" size="3"></td>
        </tr>
        <tr>
          <td >下载列表页调用条数：
            </td>
          <td ><input name="downinfo" type="text" id="downinfo" value="<%=rs("downinfo")%>" size="3"></td>
        </tr>
        <tr>
          <td >视频列表页调用条数：
            </td>
          <td ><input name="tvinfo" type="text" id="tvinfo" value="<%=rs("tvinfo")%>" size="3"></td>
        </tr>
        <tr>
          <td >招聘列表页调用条数：
          </td>
          <td ><input name="zpinfo" type="text" id="zpinfo" value="<%=rs("zpinfo")%>" size="3"></td>
        </tr>
		<tr>
		  <td >营销网点列表页显示条数：
            </td>
          <td ><input name="yxwlinfo" type="text" id="yxwlinfo" value="<%=rs("yxwlinfo")%>" size="3"></td>
	</tr>
		<tr>
          <td >留言列表页显示条数：
            </td>
          <td ><input name="lyinfo" type="text" id="lyinfo" value="<%=rs("lyinfo")%>" size="3"></td>
		</tr>
        <tr>
          <td >会员中心我的留言显示条数：
            </td>
          <td ><input name="hylyinfo" type="text" id="hylyinfo" value="<%=rs("hylyinfo")%>" size="3"></td>
        </tr>
        <tr>
          <td >会员中心我的应聘显示条数：
            </td>
          <td ><input name="hyypinfo" type="text" id="hyypinfo" value="<%=rs("hyypinfo")%>" size="3"></td>
        </tr>
        <tr>
          <td >会员中心我的订单显示条数：
            </td>
          <td ><input name="hyorderinfo" type="text" id="hyorderinfo" value="<%=rs("hyorderinfo")%>" size="3"></td>
        </tr>
		<tr>
          <td >搜索结果订单显示条数：
            </td>
          <td ><input name="sodd" type="text" id="sodd" value="<%=rs("sodd")%>" size="3"></td>
		</tr>
      <tr>
    <td colspan="2" >
      <input type="submit" name="Submit" value="保存" class="bnt" />  </td>
  </tr>
</table>
</form>

</body>
</html>

<% conn.Close
    Set conn = Nothing%>