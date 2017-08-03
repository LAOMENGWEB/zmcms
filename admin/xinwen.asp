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
table{ border:1px dotted #D4D4D4;font-size: 14px; border-right:none}
table td{BORDER-bottom: #D4D4D4 1px dotted;BORDER-right: #D4D4D4 1px dotted;padding:10px; font-family:"微软雅黑";color:#ccc; font-weight:bold;}
table td a{color:#ffffff}
</style>
<link href="css/style.css" rel="stylesheet" type="text/css">
	<script language="JavaScript" src="js/jquery-1.7.2.min.js"></script>
<style type="text/css">
.table1{ font-size: 14px; background: #F9f9f9;border:none;}
.table1 td{padding:10px; font-family:"微软雅黑";color:#ccc; font-weight:bold;}
.table1 td a{color:#ffffff}
</style>
<!--#include file="hqyy.asp"-->
</head>

<body style="background:#eeeeee">


<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center"><%=xinwen%>管理  <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>

<table border="0" align="center" cellpadding="0" cellspacing="0" style="width:390px;border:none; margin-top:20px">
  <tr>
    <td style="border:none"><table border="0" cellpadding="0" cellspacing="0"  style=" background: #28b576;width:120px;height:100px;border:1px #eeeeee solid; ">
<tr>
 <td  style="border-bottom:none;border-right:none;"><div align="center"> <a href="news_manage.asp" target="main" style="color:#ffffff"><%=xinwen%>列表管理</a></div></td>
</tr>
</table></td>
    <td style="border:none"><table border="0" cellpadding="0" cellspacing="0"  style=" background: #28b576;width:120px;height:100px;border:1px #eeeeee solid; ">
<tr>
 <td  style="border-bottom:none;border-right:none;"><div align="center"> <a href="news_add.asp" target="main" style="color:#ffffff">添加<%=xinwen%></a></div></td>
</tr>
</table></td>
    <td style="border:none"><table border="0" cellpadding="0" cellspacing="0"  style=" background: #28b576;width:120px;height:100px;border:1px #eeeeee solid; ">
<tr>
 <td  style="border-bottom:none;border-right:none;"><div align="center"> <a href="fl_news.asp" target="main" style="color:#ffffff"><%=xinwen%>分类管理</a></div></td>
</tr>
</table></td>

  </tr>
</table>





</body>
</html>
