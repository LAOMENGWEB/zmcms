<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!-- #include file="admin.asp" -->
<!-- #include file="Inc/function.asp" -->
<%
theInstalledObjects = "JMail.SMTPMail"
if Instr(session("AdminPurview"),"|91,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</script>"
  response.end
end if
'========判断是否具有管理权限
%>
<%  

maillock = request("maillock")
MailServer = request("MailServer")
MailAdd=request("MailAdd")
MailPassword=request("MailPassword")
MailUserName=request("MailUserName")
PostMail=request("PostMail")
sendmail=request("sendmail")
newguest=request("newguest")
newyp=request("newyp")
newdd=request("newdd")
hfguest=request("hfguest")
hfyp=request("hfyp")
hfdd=request("hfdd")
zchy=request("zchy")
getpass=request("getpass")

If request.Form("submit") = "保存" Then
Set rs = server.CreateObject("adodb.recordset")
sql = "select * from wzset where lang='"&lang&"' "
rs.Open sql, conn, 1, 3
 rs("maillock")= maillock 
 rs("MailServer")= MailServer
rs("MailAdd")=MailAdd
rs("MailPassword")=MailPassword
rs("MailUserName")=MailUserName
rs("PostMail")=PostMail
rs("sendmail")=sendmail
rs("newguest")=newguest
rs("newyp")=newyp
rs("newdd")=newdd
rs("hfguest")=hfguest
rs("hfyp")=hfyp
rs("hfdd")=hfdd
rs("zchy")=zchy
rs("getpass")=getpass
rs.update
    rs.Close
    Set rs = Nothing
   
	call addlog("邮件配置")
  response.Write "<script language=javascript>alert(""恭喜您修改成功"");window.location.href=""emaill.asp""</script>"
    response.End
end if
%>


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>zmcms_管理页面</title>
<link rel="stylesheet" href="css/jquery.bigcolorpicker.css" type="text/css" />
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

<form method="POST" action="emaill.asp" id="form" name="form">
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
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">邮箱配置  <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>

<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="set.asp" target="main">返回菜单</a></div></td>
 

 </tr>
</table>
<table width="99%" border="0" align="center" cellpadding="0" cellspacing="0" class="dtable4">
  
      <tr>
        <td width="19%">是否开启邮箱发送功能：</td>
		
		
		 <td width="81%">
		<select name="Maillock" >
                      <option value="0"  <%if rs("Maillock")=0 then Response.Write("selected")%>>关闭网站发送邮件功能</option>
                      <option value="1"  <%if rs("Maillock")=1 then Response.Write("selected")%> <%If Not IsObjInstalled(theInstalledObjects) Then Response.Write("disabled") %>>开启网站邮件发送功能</option>
                       
        </select></td>
      </tr>
      <tr>
        <td>新留言是否邮件提醒管理员：</td>
	
		
		 <td>
<input type="radio" name="newguest" value="1" <%if rs("newguest")="1" then response.write "checked"%>> 
是
 <input type="radio" name="newguest" value="2" <%if rs("newguest")="2" then response.write "checked"%>>
否	 
		   </td></tr>
		<tr>
		
		 <td> 新应聘是否邮件提醒管理员：</td>
	
		
		 <td><input type="radio" name="newyp" value="1" <%if rs("newyp")="1" then response.write "checked"%>> 
是
 <input type="radio" name="newyp" value="2" <%if rs("newyp")="2" then response.write "checked"%>> 
否	
			</td></tr>
		
		<tr>
		 <td>新订单是否邮件提醒管理员：</td>
	
		
		 <td><input type="radio" name="newdd" value="1" <%if rs("newdd")="1" then response.write "checked"%>> 
是
 <input type="radio" name="newdd" value="2" <%if rs("newdd")="2" then response.write "checked"%>>
否	 
			</td></tr>
		<tr>
		
		 <td>回复留言是否邮件提醒客户：</td>
	
		
		 <td><input type="radio" name="hfguest" value="1" <%if rs("hfguest")="1" then response.write "checked"%>> 
是
 <input type="radio" name="hfguest" value="2" <%if rs("hfguest")="2" then response.write "checked"%>> 
否	
			</td></tr>
		
		<tr>
		 <td>回复应聘是否邮件提醒客户：</td>
	
		
		 <td><input type="radio" name="hfyp" value="1" <%if rs("hfyp")="1" then response.write "checked"%>> 
是
 <input type="radio" name="hfyp" value="2" <%if rs("hfyp")="2" then response.write "checked"%>> 
否	 
			</td></tr>
		
		<tr>
		 <td>处理订单是否邮件提醒客户：</td>
	
		
		 <td><input type="radio" name="hfdd" value="1" <%if rs("hfdd")="1" then response.write "checked"%>> 
是
 <input type="radio" name="hfdd" value="2" <%if rs("hfdd")="2" then response.write "checked"%>>  
否	
			</td></tr>
		<tr>
		
		 <td>注册会员是否发送信件：</td>
	
		
		 <td><input type="radio" name="zchy" value="1" <%if rs("zchy")="1" then response.write "checked"%>> 
是
 <input type="radio" name="zchy" value="2" <%if rs("zchy")="2" then response.write "checked"%>> 
否        </td>
      </tr>
	  
	  
	  		<tr>
		
		 <td>找回密码是否发送信件：</td>
	
		
		 <td><input type="radio" name="getpass" value="1" <%if rs("getpass")="1" then response.write "checked"%>> 
是
 <input type="radio" name="getpass" value="2" <%if rs("getpass")="2" then response.write "checked"%>> 
否        </td>
      </tr>
	  
	  
      <tr>
        <td>smtp服务器地址：</td>
		
		
		 <td>
          
<input name="MailServer" type="text" id="MailServer" value="<%=rs("MailServer")%>" />        
例如：smtp.163.com</td>
      </tr>
     <tr>
      <td >smtp服务器信箱登录名：</td>
		
		
		 <td>
         
         
         <input name="MailAdd" type="text" id="MailAdd" value="<%=rs("MailAdd")%>" />
       注意要与发信人邮件地址一致</td>
    </tr>
    <tr>
      <td >smtp服务器信箱登录密码：</td>
		
		
		 <td>
        <input name="MailPassword" type="password" id="MailPassword" value="<%=rs("MailPassword")%>" size="40" maxlength="255"></td>
    </tr>
    <tr>
      <td >发信人姓名：</td>
		
		
		 <td>
        <input name="MailUserName" type="text" id="MailUserName" value="<%=rs("MailUserName")%>" size="40" maxlength="255">
        例如:追梦工作室      </td>
    </tr>
	<tr>
      <td >发信人邮件地址：</td>
		
		
		 <td>
        <input name="sendmail" type="text" id="sendmail" value="<%=rs("sendmail")%>" size="40" maxlength="255">
        例如：asd6583131@163.com </td>
    </tr>

   <tr>
      <td >收信邮件地址：</td>
		
		
		 <td>
        <input name="PostMail" type="text" id="PostMail" value="<%=rs("PostMail")%>" size="40" maxlength="255">
       例如：asd6583131@163.com </td>
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
<%
 conn.Close
    Set conn = Nothing
%>
