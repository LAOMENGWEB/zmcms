<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="inc/head.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="conn.asp"-->
<%
Server.ScriptTimeout=9999999
set rs1=Server.CreateObject("ADODB.Recordset")
sql="select * from wzset where lang='"&lang&"'"
rs1.open sql,conn,1,1
smtpserver =rs1("MailServer")
smtpuser =rs1("MailAdd")
smtppwd =rs1("MailPassword")
smtpemail =rs1("sendmail")
hemail=rs1("PostMail")
%>
<p>正在发送邮件<span id="sendtest"></span></p>
	  <%
Sub sendmail(mailurl)
jmail.logging = False
jmail.Silent = True
jmail.Charset = "GB2312"
jmail.ContentType = "text/html"
jmail.ReplyTo = hemail
jmail.AddRecipient mailurl
jmail.From = smtpuser
jmail.FromName= session("name2")
jmail.MailServerUserName = smtpuser
jmail.MailServerPassword = smtppwd
jmail.Subject =session("title2")
jmail.htmlBody ="<html><head><META content=zh-cn http-equiv=Content-Language><meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312""> </head><body>"&session("content2")&"</body></html>"
If jmail.send(smtpserver) Then
response.write mailurl&" 发送成功<br />"
Else
response.write mailurl&" <font color=red>发生失败</font><br />"
End If
jmail.ClearRecipients()
End Sub

%>

<%
mail=request("lmail2")
somail=Split(mail,",")
Set jmail = Server.CreateObject("JMAIL.Message")
For i = 0 To UBound(somail)
response.write "<script type=""text/javascript"">document.getElementById(""sendtest"").innerHTML="" [ " & i+1 & " / "&UBound(somail)+1&"] "";</script>"
sendmail somail(i)
Response.Flush
next
jmail.Close()
Set jmail=Nothing
%>