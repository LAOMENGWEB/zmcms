<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="admin.asp"-->
<!--#include file="conn.asp"-->
<!--#include file="inc/head.asp"-->
<%
Server.ScriptTimeout=9999999
set rs1=Server.CreateObject("ADODB.Recordset")
sql="select * from wzset where lang='"&lang&"'"
rs1.open sql,conn,1,1
smtp =rs1("MailServer")
username =rs1("MailAdd")
userpassword =rs1("MailPassword")
useremail =rs1("sendmail")
hemail=rs1("PostMail")
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
</head>
<script>
var secs = <%=session("sendtime")%>;
for(i=0;i<=secs;i++) {
   window.setTimeout("update(" + i + ")", i * 1000);
}
function update(num) {
 printnr = secs-num;
 document.getElementById('sec').innerHTML = printnr;
}
</script>
<body>
<table border="0" width="100%" id="table1" cellspacing="0" cellpadding="0">
	<tr>
	  <td align=center>
<%
set rs = Server.CreateObject("ADODB.Recordset")
sql="select emaill from demaill where lang='"&lang&"'"
rs.open sql,conn,1,1

totalcount=request("totalcount")
if isnull(totalcount) or totalcount="" then
	totalcount=0
end if

dim email,nowpage
'***********************************分页
count=cint(session("sendnum"))

if not rs.eof then
	rs.pagesize=count
	pagecount=rs.pagecount
	if request("pageid")="" then
		nowpage=1
	else
		nowpage=int(request("pageid"))
	end if
	if nowpage>rs.pagecount then
		response.write "发送完毕,服务器："&smtp&" 总共发送邮件<b><font color=red>"&totalcount&"</font></b>封"
		response.end
	end if
	if nowpage>=rs.pagecount then
		nowpage=rs.pagecount
	elseif nowpage<=1 then
		nowpage=1
	end if
	rs.absolutepage=nowpage
else
	pagecount=1
	nowpage=1
end if 
a=1
Dim JMail,msg
Set JMail = Server.CreateObject("JMail.Message") 
do while not rs.eof and a<=count


	JMail.Charset = "gb2312"
	JMail.ContentType = "text/html"
	JMail.From = useremail
	JMail.FromName = session("sendname")
	JMail.Subject = session("title") 
	jmail.ReplyTo = hemail
	JMail.MailServerUserName = username 
	JMail.MailServerPassword = userpassword
	JMail.AddHeader "Originating-IP", Request.ServerVariables("REMOTE_ADDR")
	JMail.ClearRecipients()
	for i=1 to 10
		if not rs.eof and a<=count then
			email=rs(0)
			JMail.AddRecipient(email)
			a=a+1
			rs.movenext
		end if
	next
	JMail.Body = session("content")
	msg=JMail.Send(smtp)
	
	
	
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
' 关闭并清除对象
JMail.Close()
Set JMail = Nothing

totalcount=cint(totalcount)+cint(a)-1
%>

正在发送中.....本次发送邮件<b><font color=red><%=totalcount%></font></b>封<br><br>
<span id="sec" style="color:red"></span>秒后继续发送下一批邮件</td>
</tr>
</table>
<meta http-equiv="refresh" content="<%=session("sendtime")%>;URL=email_send.asp?pageid=<%=cint(nowpage)+1%>&totalcount=<%=totalcount%>">
</body>
</html>