<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="Inc/Function.asp"-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=chanpin%></title>
<style type="text/css">
table{ border:1px dotted #D4D4D4;font-size: 14px; border-right:none}
table td{BORDER-bottom: #D4D4D4 1px dotted;BORDER-right: #D4D4D4 1px dotted;padding:10px; font-family:"微软雅黑";color:#ccc; font-weight:bold;}
table td a{color:#ffffff}
</style>
<link href="css/style.css" rel="stylesheet" type="text/css">
</head>

<body style="background:#eeeeee">
<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" style=" background: #F9f9f9;border:none; ">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">微信系统设置菜单</div></td>
</tr>
</table>

<table border="0" align="center" cellpadding="0" cellspacing="0" style="width:650px;border:none; margin-top:20px">
  <tr>
    <td style="border:none">
	<table border="0" cellpadding="0" cellspacing="0"  style=" background: #28b576;width:120px;height:100px;border:1px #eeeeee solid; ">
<tr>
 <td  style="border-bottom:none;border-right:none;"><div align="center"> <a href="weixin/set.asp" target="main" style="color:#ffffff">基本设置</a></div></td>
</tr>
</table></td>
    <td style="border:none"><table border="0" cellpadding="0" cellspacing="0"  style=" background: #28b576;width:120px;height:100px;border:1px #eeeeee solid; ">
<tr>
 <td  style="border-bottom:none;border-right:none;"><div align="center"> <a href="weixin/set2.asp" target="main" style="color:#ffffff">高级设置</a></div></td>
</tr>
</table></td>
   
<td style="border:none"><table border="0" cellpadding="0" cellspacing="0"  style=" background: #28b576;width:120px;height:100px;border:1px #eeeeee solid; ">
<tr>
 <td  style="border-bottom:none;border-right:none;"><div align="center"> <a href="weixin/KeyWord.asp" target="main" style="color:#ffffff">关键字管理</a></div></td>
</tr>
</table></td>
<td style="border:none"><table border="0" cellpadding="0" cellspacing="0"  style=" background: #28b576;width:120px;height:100px;border:1px #eeeeee solid; ">
<tr>
 <td  style="border-bottom:none;border-right:none;"><div align="center"> <a href="weixin/KeyWord.asp?act=add" target="main" style="color:#ffffff">添加关键字</a></div></td>
</tr>
</table></td>



   <td style="border:none">
	<table border="0" cellpadding="0" cellspacing="0"  style=" background: #28b576;width:120px;height:100px;border:1px #eeeeee solid; ">
<tr>
 <td  style="border-bottom:none;border-right:none;"><div align="center"> <a href="weixin/keyword_class.asp" target="main" style="color:#ffffff">关键字分类管理</a></div></td>
</tr>
</table></td>
  </tr>
</table>




<table border="0" align="center" cellpadding="0" cellspacing="0" style="width:650px;border:none; margin-top:20px">
  <tr>
    <td style="border:none"><table border="0" cellpadding="0" cellspacing="0"  style=" background: #28b576;width:120px;height:100px;border:1px #eeeeee solid; ">
<tr>
 <td  style="border-bottom:none;border-right:none;"><div align="center"> <a href="weixin/user.asp" target="main" style="color:#ffffff">粉丝管理</a></div></td>
</tr>
</table></td>
    <td style="border:none"><table border="0" cellpadding="0" cellspacing="0"  style=" background: #28b576;width:120px;height:100px;border:1px #eeeeee solid; ">
<tr>
 <td  style="border-bottom:none;border-right:none;"><div align="center"> <a href="weixin/msg.asp" target="main" style="color:#ffffff">消息管理</a></div></td>
</tr>
</table></td>
    <td style="border:none"><table border="0" cellpadding="0" cellspacing="0"  style=" background: #28b576;width:120px;height:100px;border:1px #eeeeee solid; ">
<tr>
 <td  style="border-bottom:none;border-right:none;"><div align="center"> <a href="weixin/menu.asp" target="main" style="color:#ffffff">自定义菜单管理</a></div></td>
</tr>
</table></td>
<td style="border:none"><table border="0" cellpadding="0" cellspacing="0"  style=" background: #28b576;width:120px;height:100px;border:1px #eeeeee solid; ">
<tr>
 <td  style="border-bottom:none;border-right:none;"><div align="center"> <a href="../weixin/Plug/picMsg/Admin.asp" target="main" style="color:#ffffff">自定义单图文</a></div></td>
</tr>
</table></td>
<td style="border:none"><table border="0" cellpadding="0" cellspacing="0"  style=" background: #28b576;width:120px;height:100px;border:1px #eeeeee solid; ">
<tr>
 <td  style="border-bottom:none;border-right:none;"><div align="center"> <a href="../weixin/Plug/picMsgLIST/Admin.asp" target="main" style="color:#ffffff">自定义多图文</a></div></td>
</tr>
</table></td>
  </tr>
</table>







<table border="0" align="center" cellpadding="0" cellspacing="0" style="width:650px;border:none; margin-top:20px">
  <tr>
    <td style="border:none"><table border="0" cellpadding="0" cellspacing="0"  style=" background: #28b576;width:120px;height:100px;border:1px #eeeeee solid; ">
<tr>
 <td  style="border-bottom:none;border-right:none;"><div align="center"> <a href="../weixin/Plug/LISTproduct/Admin.asp" target="main" style="color:#ffffff">系统产品素材</a></div></td>
</tr>
</table></td>
    <td style="border:none"><table border="0" cellpadding="0" cellspacing="0"  style=" background: #28b576;width:120px;height:100px;border:1px #eeeeee solid; ">
<tr>
 <td  style="border-bottom:none;border-right:none;"><div align="center"> <a href="../weixin/Plug/LISTnews/Admin.asp" target="main" style="color:#ffffff">系统新闻素材</a></div></td>
</tr>
</table></td>
    <td style="border:none"><table border="0" cellpadding="0" cellspacing="0"  style=" background: #28b576;width:120px;height:100px;border:1px #eeeeee solid; ">
<tr>
 <td  style="border-bottom:none;border-right:none;"><div align="center"> <a href="../weixin/Plug/LISTcase/Admin.asp" target="main" style="color:#ffffff">系统图片素材</a></div></td>
</tr>
</table></td>
<td style="border:none"><table border="0" cellpadding="0" cellspacing="0"  style=" background: #28b576;width:120px;height:100px;border:1px #eeeeee solid; ">
<tr>
 <td  style="border-bottom:none;border-right:none;"><div align="center"> <a href="../weixin/Plug/LISTgoods/Admin.asp" target="main" style="color:#ffffff">系统单页素材</a></div></td>
</tr>
</table></td>
 <td style="border:none"><table border="0" cellpadding="0" cellspacing="0"  style=" background: #28b576;width:120px;height:100px;border:1px #eeeeee solid; ">
<tr>
 <td  style="border-bottom:none;border-right:none;"><div align="center"> <a href="weixin/KeyWord.asp?act=edit&key=关注回复" target="main" style="color:#ffffff">关注回复设置</a></div></td>
</tr>
</table></td>
  </tr>
</table>


















</body>
</html>



























