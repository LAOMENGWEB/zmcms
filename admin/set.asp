<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="Inc/Function.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=danye%></title>
<style type="text/css">
table{ border:1px dotted #D4D4D4;font-size: 14px;  border-right:none}
table td{BORDER-bottom: #D4D4D4 1px dotted;BORDER-right: #D4D4D4 1px dotted;padding:10px; font-family:"微软雅黑";color:#ccc; font-weight:bold;}
table td a{color:#ffffff}
</style>
<link href="css/style.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="js/jquery-1.7.2.min.js"></script>
<!--#include file="hqyy.asp"-->
</head>

<body style="background:#eeeeee">
<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" style=" background: #F9f9f9;border:none; ">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">系统基本设置  <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>

<table width="421" border="0" align="center" cellpadding="0" cellspacing="0" style="width:390px;border:none; margin-top:20px">
  <tr>
    <td width="140" style="border:none"><table border="0" cellpadding="0" cellspacing="0"  style=" background: #28b576;width:120px;height:100px;border:1px #eeeeee solid; ">
<tr>
 <td  style="border-bottom:none;border-right:none;"><div align="center"> <a href="wzset.asp" target="main" style="color:#ffffff">基本设置</a></div></td>
</tr>
</table></td>
    <td width="140" style="border:none"><table border="0" cellpadding="0" cellspacing="0"  style=" background: #28b576;width:120px;height:100px;border:1px #eeeeee solid; ">
<tr>
 <td  style="border-bottom:none;border-right:none;"><div align="center"> <a href="lbsm.asp" target="main" style="color:#ffffff">列表页数量</a></div></td>
</tr>
</table></td>
    <td width="141" style="border:none"><table border="0" cellpadding="0" cellspacing="0"  style=" background: #28b576;width:120px;height:100px;border:1px #eeeeee solid; ">
<tr>
 <td  style="border-bottom:none;border-right:none;"><div align="center"> <a href="cj.asp" target="main" style="color:#ffffff">插件是否开启</a></div></td>
</tr>
</table></td>
  </tr>
</table>

<table border="0" align="center" cellpadding="0" cellspacing="0" style="width:390px;border:none; margin-top:20px">
  <tr>
    <td style="border:none"><table border="0" cellpadding="0" cellspacing="0"  style=" background: #28b576;width:120px;height:100px;border:1px #eeeeee solid; ">
<tr>
 <td  style="border-bottom:none;border-right:none;"><div align="center"> <a href="qtsz.asp" target="main" style="color:#ffffff">其他设置</a></div></td>
</tr>
</table></td>
    <td style="border:none"><table border="0" cellpadding="0" cellspacing="0"  style=" background: #28b576;width:120px;height:100px;border:1px #eeeeee solid; ">
<tr>
 <td  style="border-bottom:none;border-right:none;"><div align="center"> <a href="emaill.asp" target="main" style="color:#ffffff">邮箱配置</a></div></td>
</tr>
</table></td>
    <td style="border:none"><table border="0" cellpadding="0" cellspacing="0"  style=" background: #28b576;width:120px;height:100px;border:1px #eeeeee solid; ">
<tr>
 <td  style="border-bottom:none;border-right:none;"><div align="center"> <a href="share.asp" target="main" style="color:#ffffff">分享参数设置</a></div></td>
</tr>
</table></td>
  </tr>
</table>
<table border="0" align="center" cellpadding="0" cellspacing="0" style="width:390px;border:none; margin-top:20px">
  <tr>
    <td style="border:none"><table border="0" cellpadding="0" cellspacing="0"  style=" background: #28b576;width:120px;height:100px;border:1px #eeeeee solid; ">
<tr>
 <td  style="border-bottom:none;border-right:none;"><div align="center"> <a href="kfcs.asp" target="main" style="color:#ffffff">客服参数设置</a></div></td>
</tr>
</table></td>
    <td style="border:none"><table border="0" cellpadding="0" cellspacing="0"  style=" background: #28b576;width:120px;height:100px;border:1px #eeeeee solid; ">
<tr>
 <td  style="border-bottom:none;border-right:none;"><div align="center"> <a href="messageskin.asp" target="main" style="color:#ffffff">即时咨询样式</a></div></td>
</tr>
</table></td>
    <td style="border:none"><table border="0" cellpadding="0" cellspacing="0"  style=" background: #28b576;width:120px;height:100px;border:1px #eeeeee solid; ">
<tr>
 <td  style="border-bottom:none;border-right:none;"><div align="center"> <a href="kefu.asp" target="main" style="color:#ffffff">网站客服管理</a></div></td>
</tr>
</table></td>
  </tr>
</table>

<table border="0" align="center" cellpadding="0" cellspacing="0" style="width:390px;border:none; margin-top:20px">
  <tr>
    <td style="border:none"><table border="0" cellpadding="0" cellspacing="0"  style=" background: #28b576;width:120px;height:100px;border:1px #eeeeee solid; ">
<tr>
 <td  style="border-bottom:none;border-right:none;"><div align="center"> <a href="hight.asp" target="main" style="color:#ffffff">菜单高亮值（生成静态时高亮使用）</a></div></td>
</tr>
</table></td>

    <td style="border:none"><table border="0" cellpadding="0" cellspacing="0"  style="background: #28b576;width:120px;height:100px;border:1px #eeeeee solid; ">
<tr>
 <td  style="border-bottom:none;border-right:none;"><div align="center"><a href="logoset.asp" target="main" style="color:#ffffff">LOGO设置</a></div></td>
</tr>
</table></td>

    <td style="border:none"><table border="0" cellpadding="0" cellspacing="0"  style=" background:none;width:120px;height:100px;border:none">
<tr>
 <td  style="border-bottom:none;border-right:none;"><div align="center"></div></td>
</tr>
</table></td>


   
  </tr>
</table>


</body>
</html>


