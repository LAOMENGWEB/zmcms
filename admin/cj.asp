<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<%
if Instr(session("AdminPurview"),"|91,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</script>"
  response.end
end if
'========判断是否具有管理权限
%>
<%  
yy = request("yy")
kq = request("kq")
jt=request("jt")
kf=request("kf")
tc=request("tc")
pfad=request("pfad")
dlad=request("dlad")
ly=request("ly")
hyzc=request("hyzc")
hysh=request("hysh")
zxyp=request("zxyp")
ykgw=request("ykgw")
qq=request("qq")
JMailDisplay=request("JMailDisplay")
If request.Form("submit") = "保存" Then
Set rs = server.CreateObject("adodb.recordset")
sql = "select * from wzset where lang='"&lang&"' "
rs.Open sql, conn, 1, 3
rs("yy")=yy
rs("jt")=jt
rs("kf")=kf
rs("tc")=tc
rs("pfad")=pfad
rs("dlad")=dlad
rs("ly")=ly
rs("hyzc")=hyzc
rs("hysh")=hysh
rs("ykgw")=ykgw
rs("zxyp")=zxyp
rs("qq")=qq
rs("kq")=kq
rs("JMailDisplay")=JMailDisplay
    rs.update
    rs.Close
    Set rs = Nothing
    call addlog("插件是否开启")
  response.Write "<script language=javascript>alert(""恭喜您修改成功"");window.location.href=""cj.asp""</script>"
    response.End
End If
%>


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>zmcms_管理页面</title>
<script type="text/javascript" src="js/jquery-1.7.2.min.js"></script>
<link href="css/style.css" rel="stylesheet" type="text/css">
<style type="text/css">
table{ font-size: 14px;  border-right:none}
table td{padding:10px; font-family:"微软雅黑";color:#ccc; font-weight:bold;}
table td a{color:#ffffff}
</style>
<!--#include file="hqyy.asp"-->
</head>


<body>

<form method="POST" action="cj.asp" id="form" name="form">
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
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">插件开启设置  <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>

<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="set.asp" target="main">返回菜单</a></div></td>
 

 </tr>
</table>
<table width="99%" border="0" align="center" cellpadding="0" cellspacing="0" class="dtable4">
  
  
   <tr>
    <td width="17%" >是否开启静态页：</td>
    <td width="83%" ><input type="radio"  name="yy" value="1" <%if rs("yy")="1" then response.write "checked"%>> 
asp 
 <input type="radio" class="none"  name="yy" value="2" <%if rs("yy")="2" then response.write "checked"%>>
html	
<input type="radio" class="none"  name="yy" value="3" <%if rs("yy")="3" then response.write "checked"%>>
伪静态(需空间支持伪静态组件)</td>
   </tr>
  
     <tr>
       <td>是否开启手机自动跳转</td>
       <td><input type="radio" class="none"  name="kq" value="1" <%if rs("kq")="1" then response.write "checked"%>> 
是
 <input type="radio" class="none"  name="kq" value="2" <%if rs("kq")="2" then response.write "checked"%>>
否	</td>
     </tr>
     <tr>
    <td>是否开启会员注册：</td>
    <td><input type="radio" class="none"  name="hyzc" value="1" <%if rs("hyzc")="1" then response.write "checked"%>> 
是
 <input type="radio" class="none"  name="hyzc" value="2" <%if rs("hyzc")="2" then response.write "checked"%>>
否	</td>
    </tr>
  
     <tr>
    <td>注册会员是否审核：</td>
    <td><input type="radio" class="none"  name="hysh" value="1" <%if rs("hysh")="1" then response.write "checked"%>> 
是
 <input type="radio" class="none"  name="hysh" value="2" <%if rs("hysh")="2" then response.write "checked"%>>
否	</td>
    </tr>
     <tr>
    <td>非会员是否可以订购产品：</td>
    <td><input type="radio" class="none"  name="ykgw" value="1" <%if rs("ykgw")="1" then response.write "checked"%>> 
是
 <input type="radio" class="none"  name="ykgw" value="2" <%if rs("ykgw")="2" then response.write "checked"%>>
否	</td>
    </tr>
	 <tr>
	  <td>非会员是否可以在线应聘：        </td>
      <td><input type="radio" class="none"  name="zxyp" value="1" <%if rs("zxyp")="1" then response.write "checked"%>> 
是
 <input type="radio" class="none"  name="zxyp" value="2" <%if rs("zxyp")="2" then response.write "checked"%>>
否	</td>
    </tr>
    <tr>
<td >分享微博功能是否开启：</td>
<td ><input type="radio" class="none"  name="jt" value="1" <%if rs("jt")="1" then response.write "checked"%>>
是
 
<input type="radio" class="none"  name="jt" value="2" <%if rs("jt")="2" then response.write "checked"%>>
否<a href="share.asp" target="main"><font color="#FF0000"> 进入分享参数</font></a></td>
    </tr>
<tr> 
<td >是否显示客服QQ： </td>
<td ><input type="radio" class="none"  name="kf" value="1" <%if rs("kf")="1" then response.write "checked"%>>
是 
<input type="radio" class="none"  name="kf" value="2" <%if rs("kf")="2" then response.write "checked"%>>
否<a href="kfcs.asp" target="main"><font color="#FF0000"> 进入客服参数</font></a></td>
</tr>
<!--
<tr> 
 <td ><span >是否显示弹窗公告：</span>
   <input type="radio"  name="tc" value="1" <%if rs("tc")="1" then response.write "checked"%> />
   是 <input type="radio"  name="tc" value="2" <%if rs("tc")="2" then response.write "checked"%>>
否</td></tr>-->
<tr> 
<td >是否启用浮动客户咨询： </td>
<td ><input name="JMailDisplay" type="radio" class="none"  value="1" <%
If rs("JMailDisplay")="1" Then Response.Write("checked")%>>
        启用
        <input name="JMailDisplay" type="radio" class="none"  value="0" <%
If rs("JMailDisplay")="0" Then Response.Write("checked")%>>
        不启用<a href="messageskin.asp" target="main"><font color="#FF0000"> 进入即时咨询客服样式设置</font></a></td>
</tr>
<tr>
<td >是否显示飘浮广告：</td>
<td ><input type="radio"  name="pfad" class="none" value="1" <%if rs("pfad")="1" then response.write "checked"%>>
是 <input type="radio"  name="pfad" class="none" value="2" <%if rs("pfad")="2" then response.write "checked"%>>
 否 <a href="plad.asp" target="main"><font color="#FF0000">进入漂浮广告设置</font></a></td>
</tr>
<tr>
               <td >是否显示对联广告：                 </td>
               <td ><input type="radio" class="none"  name="dlad" value="1" <%if rs("dlad")="1" then response.write "checked"%>>
  是 <input type="radio" class="none"  name="dlad" value="2" <%if rs("dlad")="2" then response.write "checked"%>>
      否<a href="dladd.asp" target="main"><font color="#FF0000"> 进入对联广告设置</font></a></td>
</tr>
    
	<tr> 
      <td >留言是否需要审核： </td>
      <td ><input type="radio" class="none"  name="ly" value="1" <%if rs("ly")="1" then response.write "checked"%>>
  是 <input type="radio" class="none"  name="ly" value="2" <%if rs("ly")="2" then response.write "checked"%>>
      否</td>
	</tr>
    
    
	<tr> 
      <td >留言是否启用悄悄话功能： </td>
      <td ><input type="radio" class="none"  name="qq" value="1" <%if rs("qq")="1" then response.write "checked"%>>
  是 <input type="radio" class="none"  name="qq" value="2" <%if rs("qq")="2" then response.write "checked"%>>
      否</td>
	</tr>

        
        
        
        
      <tr>
    <td colspan="2" >
      <input type="submit" name="Submit" value="保存" class="bnt" /> </td>
  </tr>
</table>
</form>

</body>
</html>

