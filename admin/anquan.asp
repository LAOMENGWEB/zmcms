<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="Inc/Function.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>安全设置</title>
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
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center"> 安全设置</div></td>
</tr>
</table>

<table border="0" align="center" cellpadding="0" cellspacing="0" style="width:390px;border:none; margin-top:20px">
  <tr>
    <td style="border:none"><table border="0" cellpadding="0" cellspacing="0"  style=" background: #28b576;width:120px;height:100px;border:1px #eeeeee solid; ">
<tr>
 <td  style="border-bottom:none;border-right:none;"><div align="center"> <a href="datebf.asp?action=BackupData" target="main" style="color:#ffffff">ACCESS数据库备份</a></div></td>
</tr>
</table></td>
    <td style="border:none"><table border="0" cellpadding="0" cellspacing="0"  style=" background: #28b576;width:120px;height:100px;border:1px #eeeeee solid; ">
<tr>
 <td  style="border-bottom:none;border-right:none;"><div align="center"> <a href="datebf.asp?action=CompressData" target="main" style="color:#ffffff">ACCESS数据库压缩</a></div></td>
</tr>
</table></td>
    <td style="border:none"><table border="0" cellpadding="0" cellspacing="0"  style=" background: #28b576;width:120px;height:100px;border:1px #eeeeee solid; ">
<tr>
 <td  style="border-bottom:none;border-right:none;"><div align="center"> <a href="datebf.asp?action=ShowRestore" target="main" style="color:#ffffff">access数据库恢复</a></div></td>
</tr>
</table></td>

  </tr>
</table>

<table border="0" align="center" cellpadding="0" cellspacing="0" style="width:390px;border:none; ">
  <tr>
    <td style="border:none"><table border="0" cellpadding="0" cellspacing="0"  style=" background: #28b576;width:190px;height:100px;border:1px #eeeeee solid; ">
<tr>
 <td  style="border-bottom:none;border-right:none;"><div align="center"> <a href="admin_sql.asp?Action=config" target="main" style="color:#ffffff">SQL注入设置</a></div></td>
</tr>
</table></td>
    <td style="border:none"><table border="0" cellpadding="0" cellspacing="0"  style=" background: #28b576;width:190px;height:100px;border:1px #eeeeee solid; ">
<tr>
 <td  style="border-bottom:none;border-right:none;"><div align="center"> <a href="admin_sql.asp" target="main" style="color:#ffffff">SQL注入信息查看</a></div></td>
</tr>
</table></td>
    

  </tr>
</table>


<table border="0" align="center" cellpadding="0" cellspacing="0" style="width:390px;border:none; ">
  <tr>
    <td style="border:none"><table border="0" cellpadding="0" cellspacing="0"  style=" background: #28b576;width:190px;height:100px;border:1px #eeeeee solid; ">
<tr>
 <td  style="border-bottom:none;border-right:none;"><div align="center"> <a href="managexg.asp" target="main" style="color:#ffffff">后台目录修改</a></div></td>
</tr>
</table></td>
    <td style="border:none"><table border="0" cellpadding="0" cellspacing="0"  style=" background: #28b576;width:190px;height:100px;border:1px #eeeeee solid; ">
<tr>
 <td  style="border-bottom:none;border-right:none;"><div align="center"> <a href="sjkxg.asp" target="main" style="color:#ffffff">数据库名称修改</a></div></td>
</tr>
</table></td>
    

  </tr>
</table>

</body>
</html>
