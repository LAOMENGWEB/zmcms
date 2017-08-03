<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="Inc/Function.asp"-->


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>zmcms_管理页面</title>
<style type="text/css">
table{ border:1px dotted #D4D4D4;font-size: 12px; background-color:#F9F9F9;  border-right:none}
table td{BORDER-bottom: #D4D4D4 1px dotted;BORDER-right: #D4D4D4 1px dotted;padding:10px}
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

<body>


<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">系统信息 <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>
	
<%
Dim theInstalledObjects(17)
theInstalledObjects(0) = "MSWC.AdRotator"
theInstalledObjects(1) = "MSWC.BrowserType"
theInstalledObjects(2) = "MSWC.NextLink"
theInstalledObjects(3) = "MSWC.Tools"
theInstalledObjects(4) = "MSWC.Status"
theInstalledObjects(5) = "MSWC.Counters"
theInstalledObjects(6) = "IISSample.ContentRotator"
theInstalledObjects(7) = "IISSample.PageCounter"
theInstalledObjects(8) = "MSWC.PermissionChecker"
theInstalledObjects(9) = "Scripting.FileSystemObject"
theInstalledObjects(10) = "adodb.connection"    
theInstalledObjects(11) = "SoftArtisans.FileUp"
theInstalledObjects(12) = "SoftArtisans.FileManager"
theInstalledObjects(13) = "JMail.SMTPMail"
theInstalledObjects(14) = "CDONTS.NewMail"
theInstalledObjects(15) = "Persits.MailSender"
theInstalledObjects(16) = "LyfUpload.UploadFile"
theInstalledObjects(17) = "Persits.Upload.1"
%>
<%
set rs = Server.CreateObject("ADODB.Recordset")
sqltext="select * from Admin where userName='"+Session("userName")+"'"
rs.Open sqltext,conn,1,1
if not rs.EOF then	
	LastLoginIP=rs("LastLoginIP")	
	LastLoginTime=rs("LastLoginTime")	
end if
rs.Close
%>

<table width="99%" border="0" cellpadding="6" cellspacing="0"  align="center" style=" border: 1px solid #E6DB55;background: #FFFBCC;margin-top:5px;">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#000000">追梦工作室提醒您：为了安全考虑，请及时修改您的后台目录，管理员密码，和数据库名称！。</td>
</tr>
</table>
<table width="99%" border="0" cellpadding="6" cellspacing="0"  align="center" style="border:1px #ddd solid;margin-top:5px;margin-bottom:5px">
                                     
                                  
                                      <tr>
                                        <td width="33%"  style="border-bottom:none"><div align="center">系统信息</div></td>
                                        <td width="33%" style="border-bottom:none"><div align="center">统计信息</div></td>
                                        <td width="33%"  style="border-bottom:none;border-right:none"><div align="center">官方信息</div></td>
                                      </tr>
</table>
<table width="99%" border="0" cellpadding="0" cellspacing="0"  align="center"  style=" margin-bottom:40px">

									  <tr   style=" line-height:20px">
									    <td    width="16%"><div align="left">登陆用户：</div></td>
									    <td    width="17%"><%=Session("username")%></td>
									    <td    width="16%"><div align="left">管理员个数：</div></td>
									    <td    width="17%"><%=conn.execute("select count(*) from admin")(0)%> 个</td>
									    <td    width="16%"><div align="center">官方网址：</div></td>
                                        <td    width="17%"><a href="http://zmxt.cn" target="_blank">点此进入</a></td>
  </tr>
									  <tr   style=" line-height:20px">
									    <td    ><div align="left">登录IP：</div></td>
									    <td    ><%=Request.ServerVariables("REMOTE_ADDR")%></td>
									    <td    ><div align="left">友情链接个数：</div></td>
									    <td    ><%=conn.execute("select count(*) from yqlj")(0)%>个</td>
									    <td    ><div align="center">官方售价：</div></td>
                                        <td    > ACCESS/MSSQL:158元/188元</td>
  </tr>
									  <tr   style=" line-height:20px">
									    <td    >身份过期：</td>
									    <td    ><%=Session.timeout%> 分钟</td>
									    <td    >新闻条数：</td>
									    <td    ><%=Conn.Execute("Select count(*) from news ")(0)%> 条</td>
									    <td    ><div align="center">售前咨询： </div></td>
                                        <td    >1666719728</td>
  </tr>
									  
									  
									  <tr>
									    <td    >上线时间：</td>
									    <td    ><%=LastLoginTime%></td>
									    <td    >产品个数：</td>
									    <td    ><%=conn.execute("select count(*) from product")(0)%> 个</td>
									    <td    ><div align="center">技术支持： </div></td>
                                        <td    >2685863740</td>
  </tr>
									  <tr>
									    <td    >服务器域名：</td>
									    <td    ><%=Request.ServerVariables("server_name")%> / <%=Request.ServerVariables("Http_HOST")%></td>
									    <td    >新简历/全部简历：</td>
									    <td    ><%=conn.execute("select count(*) from HrDemandAccept where replycontent is null")(0)%>/<%=conn.execute("select count(*) from HrDemandAccept")(0)%> [<a href="yp.asp?sh=3">查看</a>]</td>
									    <td    ><div align="center">程序购买： </div></td>
                                        <td    >2685863740</td>
  </tr>
									  <tr>
									    <td    >脚本解释引擎：</td>
									    <td    ><%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
									    <td    >图片个数：</td>
									    <td    ><%=conn.execute("select count(*) from al")(0)%> 个</td>
									    <td    ><div align="center">程序制作：</div></td>
                                        <td    ><a href="http://zmxt.cn">追梦工作室</a></td>
  </tr>
									  <tr>
									    <td    >服务器软件的名称：</td>
									    <td    ><%=Request.ServerVariables("SERVER_SOFTWARE")%></td>
									    <td    >下载个数：</td>
									    <td    ><%=conn.execute("select count(*) from download")(0)%> 个</td>
									    <td    ><div align="center">帮助文档：</div></td>
                                        <td    ><a href="http://www.zmxt.cn/help">点此进入</a></td>
  </tr>
									  <tr>
									    <td    >FSO文本读写：     </td>
									     <td    > <%If Not IsObjInstalled(theInstalledObjects(9)) Then%>
      <b>×</b> 
      <%else%>
      <b>√</b> 
      <%end if%></td>
									     <td    >视频个数：</td>
									     <td    ><%=conn.execute("select count(*) from zxsp")(0)%> 个</td>
								        <td    ><div align="center">更新时间：</div></td>
								        <td    >20150908</td>
  </tr>
									  <tr>
									    <td    >数据库使用：      </td>
									    <td    ><%If Not IsObjInstalled(theInstalledObjects(10)) Then%>
      <b>×</b> 
      <%else%>
      <b>√</b> 
      <%end if%></td>
									    <td    >新订单/全部订单：</td>
									    <td    ><%=conn.execute("select count(*) from orderdd where isdate=0")(0)%>/<%=conn.execute("select count(*) from orderdd ")(0)%> [<a href="order.asp?sh=2">查看</a>]</td>
									    <td colspan="2"    >&nbsp;</td>
  </tr>
									  <tr>
									    <td  style=" border-bottom:none"  >Jmail组件支持：      </td>
									    <td  style=" border-bottom:none"  ><%If Not IsObjInstalled(theInstalledObjects(13)) Then%> <b>×</b> <%else%> <b>√</b> <%end if%> </td>
									    <td colspan="2"    style=" border-bottom:none">&nbsp;</td>
									    <td colspan="2"    style=" border-bottom:none">&nbsp;</td>
  </tr>
   </table>
</body>
</html>
