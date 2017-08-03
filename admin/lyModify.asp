<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<% 
If request.Form("submit") = "回复" Then 
id=request("id")
hf=Trim(request.form("hf"))
sh=Trim(request.form("sh"))
if id="" or not isnumeric(id) then
Response.Write "<script>alert('参数错误！');history.go(-1);</script>" 
Response.End()
end if

Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from ly where id="&id
rs.open sql,conn,1,3
If rs.EOF Then
response.Write "<script language=javascript>alert(""不存在此数据！"");window.history.back();</script>"
response.End
Else	
rs("hf") = hf
rs("ReplyTime") = now()
if sh="1" then
		rs("sh")=1
	else
		rs("sh")=0
	end if
rs.update
mtitle=rs("title")
mcontent=rs("content")
memaill=rs("emaill")
mphone=rs("phone")
mqq=rs("qq")
msj=rs("sj")
mReplyTime=rs("ReplyTime")
mhf=rs("hf")
rs.close
set rs=nothing
if maillock=1 and hfguest=1 then
SendStat = Jmail(""&memaill&"",""&word(625)&""&msj&""&word(626)&""&mc&""&word(627)&"","<br>"&word(628)&""&mcontent&",<br>"&word(629)&""&mhf&"<br><br>==============================<br><br>"&word(630)&""&mc&""&word(631)&""&mReplyTime&"<br>"&mc&"<br>"&word(632)&""&ym&"<br>"&word(633)&""&yx&"<br>"&word(634)&""&sj&"<br>"&word(635)&""&cz&"<br>"&word(636)&""&lxdz&"<br>==============================","GB2312","text/html") 
end if
call addlog("回复新留言")
if maillock=1 and hfguest=1 then
response.Write "<script language=javascript>alert(""恭喜您回复成功,留言信息已经发送邮件到对方的邮箱内！"");window.location.href=""ly.asp""</script>"
else
response.Write "<script language=javascript>alert(""恭喜您回复成功"");window.location.href=""ly.asp""</script>"
end if
 response.End
end if
end if
%>
<%
id=request.querystring("id")
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "Select * From ly where id="&id&" and lang='"&lang&"'", conn,1,1

If rs.EOF Then
    response.Write "<script language=javascript>window.location.href=""ly.asp""</script>"
    response.End
End if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>zmcms_管理页面</title>
<script language="JavaScript" src="js/jquery-1.7.2.min.js"></script>
 <script type="text/javascript" src="../js/Validform_v5.3.2_min.js"></script>
 <link rel="stylesheet" href="../zmbjq/themes/default/default.css" />
<link rel="stylesheet" href="../zmbjq/plugins/code/prettify.css" />
<script charset="utf-8" src="../zmbjq/kindeditor-all.js"></script>
<script charset="utf-8" src="../zmbjq/lang/zh_CN.js"></script>
<script charset="utf-8" src="../zmbjq/plugins/code/prettify.js"></script>
 <link href="css/style.css" rel="stylesheet" type="text/css">
  <script type="text/javascript">
$(function(){
	
$(".checkform").Validform();  //就这一行代码！;
	
})
</script>
		
	   <script>
 var editor;
 KindEditor.ready(function(K) {
 editor = K.create('textarea[name="hf"]', {
 filterMode : false, 
 uploadJson : '../zmbjq/asp/upload_json.asp',
 fileManagerJson : '../zmbjq/asp/file_manager_json.asp',
 allowFileManager : true,
 width: '700px',
 hight: '300px',
 afterBlur: function(){this.sync();}
 
 });
 	K('input[name=appendHtml]').click(function(e) {
					editor.insertHtml('[zmcmsfy]');
				});
			
			prettyPrint();
 });
 </script>
 <style type="text/css">
.table1{ font-size: 14px; background: #F9f9f9;border:none;}
.table1 td{padding:10px; font-family:"微软雅黑";color:#ccc; font-weight:bold;}
.table1 td a{color:#ffffff}
</style>
<!--#include file="hqyy.asp"-->
</head>

<body>
<form method="post" name="myform" class="checkform">

<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">回复留言 <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>




<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="ly.asp" target="main">留言管理</a></div></td>
 

 </tr>
</table>






  <table width="99%"  border="0" cellpadding="0" cellspacing="0" align="center" class="dtable4">
				  
				  <tr >
                      <td width="17%"  >用户：</td>
                      <td width="83%"><%=rs("title")%></td>
    </tr>
                  
					<tr >
                      <td  >QQ：</td><td><%=rs("qq")%></td>
                    </tr>
					<tr >
                      <td  >邮箱：</td><td><%=rs("emaill")%></td>
                    </tr>
						<tr >
						  <td  >电话：</td><td><%=rs("phone")%></td>
				    </tr>
						<tr >
                      <td  >留言内容：</td><td><%=rs("content")%></td>
                    </tr>
                  
                        <tr >
                          <td  >回复内容：</td><td>
                          <textarea name="hf" style="width:80%;height:200px;visibility:hidden;" datatype="*" nullmsg="回复内容不能为空"><%=rs("hf")%></textarea></td>
                        </tr>
                        <tr >
                          <td  >是否审核通过并在前台显示：</td><td>
                          <input name="sh" class="none" type="checkbox" id="sh" value="1" <% if rs("sh")=1 then response.Write("checked") end if%>></td>
                        </tr>
                    <tr >
                      <td   colspan="2">
                        <div align="left">
                          <input name="submit" type="submit" value="回复" class="bnt">
                          &nbsp;</div></td></tr>
  </table>
             
</form>

<%
Function   Jmail(mailTo,mailTopic,mailBody,mailCharset,mailContentType)   
  
          Dim   ConstFromNameCn,ConstFromNameEn,ConstFrom,ConstMailDomain,ConstMailServerUserName,ConstMailServerPassword   
    
          ConstFromNameCn   =   ""&MailUserName&""   
          ConstFromNameEn   =   ""&MailUserNameen&""   
          ConstFrom   =   ""&sendmail&""   
          ConstMailDomain   =   ""&MailServer&""   
          ConstMailServerUserName   =   ""&MailAdd&""   
          ConstMailServerPassword   =   ""&MailPassword&""'   
  
          On   Error   Resume   Next   
          Dim   myJmail   
          Set   myJmail   =   Server.CreateObject("JMail.Message")   
          myJmail.Logging   =   False 
          myJmail.ISOEncodeHeaders   =   False 
          myJmail.ContentTransferEncoding   =   "base64"   
          myJmail.AddHeader   "Priority","3"   
          myJmail.AddHeader   "MSMail-Priority","Normal"   
          myJmail.AddHeader   "Mailer","Microsoft   Outlook   Express   6.00.2800.1437"   
          myJmail.AddHeader   "MimeOLE","Produced   By   Microsoft   MimeOLE   V6.00.2800.1441"   
          myJmail.Charset   =   mailCharset   
          myJmail.ContentType   =   mailContentType   
    
          If   UCase(mailCharset)   =   "GB2312"   Then   
                  myJmail.FromName   =   ConstFromNameCn   
          Else   
                  myJmail.FromName   =   ConstFromNameEn   
          End   If   
    
          myJmail.From   =   ConstFrom   
          myJmail.Subject   =   mailTopic   
          myJmail.Body   =   mailBody   
          myJmail.AddRecipient   mailTo   
          myJmail.MailDomain   =   ConstMailDomain   
          myJmail.MailServerUserName   =   ConstMailServerUserName   
          myJmail.MailServerPassword   =   ConstMailServerPassword   
          myJmail.Send   ConstMailDomain   
          myJmail.Close   
          Set   myJmail=nothing     
    
          If   Err   Then     
                  Jmail=Err.Description   
                  Err.Clear   
          Else   
                  Jmail="OK"   
          End   If   
    
          On   Error   Goto   0   
End   Function  %>
</body>
</html>
