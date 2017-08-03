<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="class_page.asp"-->
<!--#include file="htmlconfig.asp"-->
<!--#include file="admin_html_function.asp"-->
<% If request.Form("submit") = "回复" Then %>
<%
id=request("id")
replycontent=Trim(request.form("replycontent"))
if id="" or not isnumeric(id) then
Response.Write "<script>alert('参数错误！');history.go(-1);</script>" 
Response.End()
end if

Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from HrDemandAccept where id="&id
rs.open sql,conn,1,3


	
rs("replycontent") = replycontent
rs("ReplyTime") = now()

rs.update
mreplycontent=rs("replycontent")
mReplyTime=rs("ReplyTime")
mEmail=rs("Email")
madddate=rs("adddate")
mQuarters=rs("Quarters")
rs.close
set rs=nothing

	if maillock=1 and hfyp=1 then
SendStat =   Jmail(""&mEmail&"",""&word(637)&""&madddate&""&word(638)&""&mQuarters&""&word(639)&"",""&word(640)&""&madddate&""&word(641)&""&mQuarters&""&word(642)&"<br><br>"&mreplycontent&"<br><br>"&word(630)&""&mc&""&word(643)&""&mReplyTime&"<br><br>==============================<br>"&mc&"<br>"&word(632)&""&ym&"<br>"&word(633)&""&yx&"<br>"&word(634)&""&sj&"<br>"&word(635)&""&cz&"<br>"&word(636)&""&lxdz&"<br>==============================","GB2312","text/html") 
end if
call addlog("回复应聘信息")

if maillock=1 and hfyp=1 then
response.Write "<script language=javascript>alert(""恭喜您回复成功,回复结果已经发送邮件到对方的邮箱！"");window.location.href=""yp.asp""</script>"
else
response.Write "<script language=javascript>alert(""恭喜您回复成功"");window.location.href=""yp.asp""</script>"
end if
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

<style type="text/css">
.table1{ font-size: 14px; background: #F9f9f9;border:none;}
.table1 td{padding:10px; font-family:"微软雅黑";color:#ccc; font-weight:bold;}
.table1 td a{color:#ffffff}
</style>
<!--#include file="hqyy.asp"-->
</head>

<body>
<%
id=request.QueryString("id")
set rs=server.createobject("adodb.recordset")
sql="select * from HrDemandAccept where id="&id&" and lang='"&lang&"' order by ID desc"
rs.open sql,conn,1,1
	If rs.EOF Then
    response.Write "<script language=javascript>window.location.href=""yp.asp""</script>"
    response.End
    End if				 
%>

<form method="post" name="myform" class="checkform">

<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">应聘详细信息  <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>



<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="zhaopin.asp" target="main">返回菜单</a></div>
 
 
 <div class="dmeun"><a href="yp.asp" target="main">应聘管理</a></div>
 </td>
 

 </tr>
</table>

  <table width="99%" border="0" cellpadding="6" cellspacing="0" align="center" class="dtable4">
                
             
<tr> 
                  <td width="17%"  > <div align="center">应聘岗位</div></td>
                  <td >&nbsp;&nbsp;<font color="red"><b><%=rs("Quarters")%></b> </font>&nbsp; <a href="yp.asp?id=<%=rs("ID")%>&amp;action=Del" onClick="return ConfirmDel();"></a>                 </td>
                  <td ><div align="center">电 话</div></td>
                  <td style="border-right:none;"><%=rs("phone")%></td>
                </tr>
                <%if rs("email")<>"" then%>
                <tr> 
                  <td height="10" > <div align="center">姓 名</div></td>
                  <td width="34%" >&nbsp;&nbsp;<%=rs("Name")%></td>
                  <td width="13%" > <div align="center">性 别</div></td>
                  <td width="36%" style="border-right:none;">&nbsp;&nbsp;<%=rs("sex")%></td>
                </tr>
                <tr> 
                  <td height="22" > <div align="center">出生日期</div></td>
                  <td >&nbsp;&nbsp;<%=rs("birthday")%></td>
                  <td > <div align="center">婚姻状况</div></td>
                  <td style="border-right:none;">&nbsp;&nbsp;<%=rs("marry")%></td>
                </tr>
                <tr> 
                  <td height="22" ><div align="center">身高</div></td>
                  <td >&nbsp;&nbsp;<%=rs("Stature")%>cm</td>
                  <td ><div align="center">户藉地</div></td>
                  <td style="border-right:none;">&nbsp;&nbsp;<%=rs("Residence")%></td>
                </tr>
                <tr> 
                  <td height="22" > <div align="center">毕业院校</div></td>
                  <td >&nbsp;&nbsp;<%=rs("school")%></td>
                  <td ><div align="center">联系地址：</div></td>
                  <td style="border-right:none;">&nbsp;&nbsp;<%=rs("add")%></td>
                </tr>
                <tr> 
                  <td height="22" > <div align="center">学 历</div></td>
                  <td >&nbsp;&nbsp;<%=rs("studydegree")%></td>
                  <td > <div align="center">专 业</div></td>
                  <td style="border-right:none;">&nbsp;&nbsp;<%=rs("specialty")%></td>
                </tr>
                <tr> 
                  <td height="22" > <div align="center">毕业时间</div></td>
                  <td >&nbsp;&nbsp;<%=rs("gradyear")%></td>
                  <td > <div align="center"><font color="#FF0000">应聘日期</font></div></td>
                  <td style="border-right:none;"><font color="#FF0000">&nbsp;&nbsp;<%=rs("Adddate")%></font></td>
                </tr>
                
                
                <tr> 
                  <td height="22" ><div align="center">邮编</div></td>
                  <td >&nbsp;&nbsp;<%=rs("Postcode")%></td>
                  <td ><div align="center">E-mail</div></td>
                  <td >&nbsp;&nbsp;<%=rs("email")%></td>
                </tr>
                <%end if%>
                <tr> 
                  <td height="25" > <div align="center">教育经历</div></td>
                  <td height="25" colspan="3" style="border-right:none;">&nbsp;&nbsp;<%=rs("Edulevel")%></td>
                </tr>
               
                <tr> 
                  <td height="25" > <div align="center">个人简历</div></td>
                  <td height="25" colspan="3" style="border-right:none;">&nbsp;&nbsp;<%=rs("Experience")%></td>
                </tr>
				  <script>
 var editor1;
 KindEditor.ready(function(K) {
 editor1 = K.create('textarea[name="replycontent"]', {
 filterMode : false, 
 uploadJson : '../zmbjq/asp/upload_json.asp',
 fileManagerJson : '../zmbjq/asp/file_manager_json.asp',
 allowFileManager : true,
 width: '700px',
 hight: '300px',
 afterBlur: function(){this.sync();}
 });

 });
 </script>
				 <tr>
                  <td height="25"  ><div align="center">回复内容</div></td>
                  <td height="25" colspan="3" style="padding-left:8px;border-right:none"><textarea name="replycontent" style="width:80%;height:200px;visibility:hidden;" datatype="*" nullmsg="回复内容不能为空"><%=rs("replycontent")%></textarea></td>
			    </tr>
				 <tr>
                 
                  <td height="25" colspan="4" style="border-right:none;border-bottom:none"><label>
                    <input type="submit" name="Submit" value="回复" class="bnt" />
                  </label></td>
			    </tr>
            </table>			  
 </form>
	<%

RS.Close
Set RS = Nothing

%>
</body>
</html>
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