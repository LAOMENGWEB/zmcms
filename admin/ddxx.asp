<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="Inc/Function.asp"-->
<!--#include file="class_page.asp"-->

<%
if Instr(session("AdminPurview"),"|23,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</script>"
  response.end
end if
'========判断是否具有管理权限

If request.Form("submitSaveEdit") = "保存" Then
id=request.QueryString("id")
replycontent=Trim(request.form("replycontent"))
if id="" or not isnumeric(id) then
Response.Write "<script>alert('参数错误！');history.go(-1);</script>" 
Response.End
end if

Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from orderdd where id="&id
rs.open sql,conn,1,3
If rs.EOF Then
    response.Write "<script language=javascript>alert(""不存在此数据！"");window.history.back();</script>"
    response.End
Else

	
rs("replycontent") = replycontent
rs("ReplyTime") = now()
rs.update
Email=rs("Email")
ReplyContent=rs("ReplyContent")
ReplyTime=rs("ReplyTime")
ddcp=rs("ddcp")
Addtime=rs("Addtime")
rs.close
set rs=nothing
if maillock=1 and hfdd=1 then
SendStat =   Jmail(""&Email&"",""&word(644)&""&Addtime&""&word(645)&"",""&word(646)&""&Addtime&""&word(647)&"<br><br>"&Replycontent&"<br><br>"&word(630)&""&mc&""&word(631)&""&Replytime&"<br><br>==============================<br>"&mc&"<br>"&word(632)&""&ym&"<br>"&word(633)&""&yx&"<br>"&word(634)&""&sj&"<br>"&word(635)&""&cz&"<br>"&word(636)&""&lxdz&"<br>==============================","GB2312","text/html") 
end if

	
	
	call addlog("处理订单信息")
	if maillock=1 and hfyp=1 then
     response.Write "<script language=javascript>alert(""成功回复此订单信息,回复结果已经发送邮件到对方的邮箱！"");window.location.href=""order.asp""</script>"
	 else
	 response.Write "<script language=javascript>alert(""成功回复此订单信息"");window.location.href=""order.asp""</script>"
	 end if
	 
	 end if
	 end if

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
<script type="text/javascript">
$(function() {
	setTimeout("changeDivStyle();", 100); // 0.1秒后展示结果，因为HTML加载顺序，先加载脚本+样式，再加载body内容。所以加个延时
});
function changeDivStyle(){
	var o_status = $("#o_status").val();	//获取隐藏框值

	if(o_status==0){
		$('#create').css('background', '#DD0000');
		$('#createText').css('color', '#DD0000');
	}else if(o_status==1){
		$('#check').css('background', '#DD0000');
		$('#checkText').css('color', '#DD0000');
	}else if(o_status==2){
		$('#produce').css('background', '#DD0000');
		$('#produceText').css('color', '#DD0000');
	}else if(o_status==3){
		$('#delivery').css('background', '#DD0000');
		$('#deliveryText').css('color', '#DD0000');
	}else if(o_status>=4){
		$('#received').css('background', '#DD0000');
		$('#receivedText').css('color', '#DD0000');
	}
}
</script>

<style type="text/css">
*{PADDING:0}
a,img{border:0;}

body{background:#ffffff;font:12px/180% Arial, Helvetica, sans-serif, "新宋体";}

.stepInfo{position:relative;margin:20px auto 0 auto;width:500px;padding-bottom:20px}
.stepInfo li{float:left;width:48%;height:0.15em;background:#bbb;}
.stepIco{border-radius:1em;padding:0.03em;background:#bbb;text-align:center;line-height:1.5em;color:#fff; position:absolute;width:1.4em;height:1.4em;}
.stepIco1{top:-0.7em;left:-1%;}
.stepIco2{top:-0.7em;left:21%;}
.stepIco3{top:-0.7em;left:46%;}
.stepIco4{top:-0.7em;left:71%;}
.stepIco5{top:-0.7em;left:95%;}
.stepText{color:#666;margin-top:0.2em;width:4em;text-align:center;margin-left:-1.4em;}
</style>
<style type="text/css">
.table1{ font-size: 14px; background: #F9f9f9;border:none;}
.table1 td{padding:10px; font-family:"微软雅黑";color:#ccc; font-weight:bold;}
.table1 td a{color:#ffffff}
</style>
<!--#include file="hqyy.asp"-->
</head>




	
</head>

<body>
	<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">订单详情  <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>
<% 
id=request.QueryString("id")
set rs=server.createobject("adodb.recordset")
sql="select * from orderdd where id="&id&" order by ID desc"
rs.open sql,conn,1,1
OrderName=rs("OrderName")
	  Remark=ReStrReplace(rs("Remark"))
	  mLinkman=GuestInfo(rs("MemID"),rs("Linkman"),rs("Sex"))
	  mCompany=rs("Company")
	  mAddress=rs("Address")
	  mZipCode=rs("ZipCode")
	  orderddh=rs("orderddh")
	  mFax=rs("Fax")
	  mMobile=rs("Mobile")
	  mEmail=rs("Email")
	  mAddTime=rs("AddTime")
	  ReplyContent=ReStrReplace(rs("ReplyContent"))
	  ReplyTime=rs("ReplyTime")				 
      shipmentmoney=rs("shipmentmoney")
	  orderjd=rs("isdate")
set rs1=server.createobject("adodb.recordset")
sql1="select * from ordercp where orderddh='"&rs("orderddh")&"' order by ID desc"
rs1.open sql1,conn,1,1
If rs1.EOF Then
    response.Write "<script language=javascript>window.location.href=""order.asp""</script>"
    response.End
    else
Quatity=rs1("bookcount")
End if
%>
<input type="hidden" value="<%=orderjd%>" id="o_status" />


<div class="stepInfo">
	<ul>
		<li></li>
		<li></li>
	</ul>
	<div class="stepIco stepIco1" id="create">1
		<div class="stepText" id="createText">创建订单等待付款</div>
	</div>
	<div class="stepIco stepIco2" id="check">2
		<div class="stepText" id="checkText">已支付</div>
	</div>
	<div class="stepIco stepIco3" id="produce">3
		<div class="stepText" id="produceText">处理中</div>
	</div>
	<div class="stepIco stepIco4" id="delivery">4
		<div class="stepText" id="deliveryText">已发货</div>
	</div>
	<div class="stepIco stepIco5" id="received">5
		<div class="stepText" id="receivedText">确认收货订单完成</div>
	</div>
</div>
<table width="99%" border="0" align="center" cellpadding="0" cellspacing="0" class="dtable2" style="margin-top:40px">
            <tr class="wz">
              <td width="12%" ><div align="center">产品订单号</div></td>
              <td width="19%" ><div align="center">产品名称</div></td>
              <td width="11%" ><div align="center">购买</div></td>
			  <td width="12%" ><div align="center">购买数量</div></td>
			  <td width="11%" style="border-right:none"><div align="center">商品品总价</div></td>
            </tr>
<%
while not rs1.eof


ksum=ccur(rs1("hyj")) * Quatity

Sum = Sum + ksum

 %>
          
		
			    <tr class="gwcsj" height="40">
			      <td ><div align="center"><%=rs1("orderddh")%></div></td> 
              <td ><div align="center"><%=rs1("productname")%>
			  
			  <%if rs1("p_id")<>0 then%>
              (
			  <%
			  set rs3=server.createobject("adodb.recordset")
              sql3="select * from product_stock where id="&rs1("p_id")&" "
              rs3.open sql3,conn,1,1
			  
			  response.write(""&rs3("s_value")&"")
			  
			  %>
                
                
                
              )
			  
			  <%end if%></div></td>
              <td ><div align="center"><%=rs1("hyj")%></div></td>
			  <td ><div align="center"><%=rs1("bookcount")%></div></td>
			  <td><div align="center"><%=formatnumber2(ksum)%></div></td>
            </tr>
           <%

      rs1.movenext
      wend
	 
	  %>

          
	    <tr>
		
	      <td colspan="5" style=" text-align:right">	  　配送费用：<%=formatnumber2(shipmentmoney)%>元 共计：<b><%=formatnumber2(sum+shipmentmoney)%></b>
          元　</td>
    </tr>
</table>

	<%
	
     rs1.close
    set rs1=nothing

  %>






  <form name="editForm" method="post" action="ddxx.asp?Action=SaveEdit&Result=<%=Result%>&ID=<%=ID%>" class="checkform">
  <table width="99%" border="0" align="center" cellpadding="0" cellspacing="0" class="dtable1">
  <tr>
    <td>订购人详细信息</td>
  </tr>
</table>
 <table width="99%" border="0" align="center" cellpadding="0" cellspacing="0" class="dtable4">

      <tr>
        <td width="10%">订单标题：</td>
        <td width="90%">
        <input name="OrderName" type="text"  id="OrderName" style="WIDTH: 240;" value="<%=OrderName%>" readonly></td>
      </tr>
	   
      <tr>
        <td>备注说明：</td><td>
      <textarea name="Remark" rows="6"  id="Remark" style="WIDTH: 76%;" readonly><%=Remark%></textarea>      </tr>
      <tr>
        <td >订购者：</td><td><%=mLinkman%></td>
      </tr >
      <tr>
        <td >单位名称：</td><td>
        <input name="Company" type="text"  id="Company" style="WIDTH: 240;" value="<%=mCompany%>" readonly></td>
      </tr>
      <tr>
        <td >通信地址：</td><td>
        <input name="Address" type="text"  id="Address" style="WIDTH: 240;" value="<%=mAddress%>" readonly></td>
      </tr>
      <tr>
        <td >邮　　编：</td><td> <input name="ZipCode" type="text"  id="ZipCode" style="WIDTH: 120" value="<%=mZipCode%>" readonly></td>
      </tr>
      
      <tr>
        <td >传　　真：</td><td>
        <input name="Fax" type="text"  id="Fax" style="WIDTH: 120" value="<%=mFax%>" readonly></td>
      </tr>
      <tr>
        <td >移动电话：</td><td>
        <input name="Mobile" type="text"  id="Mobile" style="WIDTH: 120" value="<%=mMobile%>" readonly /></td>
      </tr>
      <tr>
        <td >电子邮箱：</td><td>
        <input name="Email" type="text"  id="Email" style="WIDTH: 240" value="<%=mEmail%>" readonly></td>
      </tr>
      <tr>
        <td >订购时间：</td><td>
        <input name="AddTime" type="text"  id="AddTime" style="WIDTH: 240" value="<%=mAddTime%>" readonly></td>
      </tr>
        <script>
 var editor1;
 KindEditor.ready(function(K) {
 editor1 = K.create('textarea[name="ReplyContent"]', {
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
        <td >回复内容：</td><td>
     
		
		<textarea name="ReplyContent" style="width:700px;height:200px;visibility:hidden;" datatype="*" nullmsg="回复内容不能为空"><%=ReplyContent%></textarea>		</td>
      </tr>
      <tr>
        <td colspan="2">
          
         
            <input name="submitSaveEdit" type="submit" class="bnt"  id="submitSaveEdit" value="保存"  >          </td>
      </tr>
    </table>
  </form>
	<%
	
     rs.close
    set rs=nothing

  %>
</BODY>
</HTML>
<%

function GuestInfo(ID,Guest,Sex)
  Dim rs,sql
  Set rs=server.CreateObject("adodb.recordset")
  sql="Select * From Members where ID="&ID
  rs.open sql,conn,1,1
  if rs.bof and rs.eof then
    GuestInfo=Guest & "&nbsp;" & Sex
  else
    GuestInfo="<font color='green'>会员&nbsp;</font><a href='MemEdit.asp?Result=Modify&ID="&ID&"' onClick='changeAdminFlag(""前台会员资料"")'>"&Guest&"</a>"&Sex
  end if
  rs.close
  set rs=nothing
end function 
%>
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