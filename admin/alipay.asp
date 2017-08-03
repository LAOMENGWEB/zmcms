<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!-- #include file="admin.asp" -->
<!-- #include file="Inc/function.asp" -->


<%
if Instr(session("AdminPurview"),"|515,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</script>"
  response.end
end if
'========判断是否具有管理权限
%>




<%  



If request.Form("submit") = "保存" Then
alipayemaill = request("alipayemaill")
alipaypid = request("alipaypid")
alipaykey=request("alipaykey")
alipayapi=request("alipayapi")

Set rs = server.CreateObject("adodb.recordset")
sql = "select * from wzset where lang='"&lang&"' "
rs.Open sql, conn, 1, 3
rs("alipayemaill")= alipayemaill 
rs("alipaypid")= alipaypid
rs("alipaykey")=alipaykey
rs("alipayapi")=alipayapi
rs("lang")=lang
rs.update
    rs.Close
    Set rs = Nothing
   
	call addlog("设置支付宝信息")
  response.Write "<script language=javascript>alert(""设置支付宝信息成功"");window.location.href=""alipay.asp""</script>"
    response.End
End If
%>


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>zmcms_管理页面</title>

<script language="JavaScript" src="js/jquery-1.7.2.min.js"></script>
 <script type="text/javascript" src="../js/Validform_v5.3.2_min.js"></script>
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

<form name="editForm" method="POST" action="alipay.asp"  >
<%
Set rs = server.CreateObject("adodb.recordset")
sql = "select * from wzset where lang='"&lang&"' "
rs.Open sql, conn, 1, 1
If rs.EOF Then
response.Write "<script language=javascript>window.location.href=""alipay.asp""</script>"
response.End
End If


%>
		

<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center"> 支付宝设置 <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>



<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td >
 <div class="dmeun"><a href="chajian.asp" target="main">返回菜单</a></div>
 
 
 </td>
 

 </tr>
</table>

<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td >

 <div class="dmeun" ><a href="alipay.asp" target="main" >支付宝设置</a></div>
 <div class="dmeun" style=" background:#eeeeee; color:#999999"><a href="blank.asp" target="main" style=" color:#999999">银行转账设置</a></div>
 
 
 </td>
 

 </tr>
</table>
<table width="99%" border="0" align="center" cellpadding="0" cellspacing="0"  class="dtable4">
  <tr>
    <td width="13%">支付宝接口：</td>
      <td width="87%"><input style="border:none" type="radio" name="alipayapi" value="1" <%if rs("alipayapi")="1" then response.write "checked"%> >
即时到帐接口&nbsp;</span>&nbsp;&nbsp;&nbsp; 
<input style="border:none" type="radio" name="alipayapi" value="2" <%if rs("alipayapi")="2" then response.write "checked"%>>
担保交易接口	&nbsp;&nbsp;&nbsp;
<input style="border:none" type="radio" name="alipayapi" value="3" <%if rs("alipayapi")="3" then response.write "checked"%>>
标准双接口</td>
  </tr>

   
 
  
      <tr>
        <td>收款支付宝账号：</td>
      <td>
          <label>
        <input name="alipayemaill" type="text" id="alipayemaill" size="30" value="<%=rs("alipayemaill")%>" />
          </label>请填写真实有效的支付宝账号，用于收取用户以现金兑换交易积分的相关款项。如您没有支付宝帐号，请点击<a href="https://lab.alipay.com/user/reg/index.htm" target="_blank">这里</a>注册</td>
      </tr>
	  
	  <tr>
        <td>合作者身份(PID)：</td>
      <td>
          <label>
        <input name="alipaypid" type="text" id="mc" size="30" value="<%=rs("alipaypid")%>" />
          合作身份者ID，以2088开头的16位纯数字</label></td>
      </tr>

      <tr>
        <td>交易安全校验码 (key):</td>
      <td>
      <input name="alipaykey" type="text" id="alipaykey" value="<%=rs("alipaykey")%>" size="40" maxlength="50">
      安全检验码，以数字和字母组成的32位字 </td>
      </tr>
     
    
      <tr>
        <td colspan="2" style=" border-bottom:none"> 提示：如何获取安全校验码和合作身份者ID<br />
          ' 1.用您的签约支付宝账号登录支付宝网站(www.alipay.com)<br />
          ' 2.点击“商家服务”(https://b.alipay.com/order/myOrder.htm)<br />
        ' 3.点击“查询合作者身份(PID)”、“查询安全校验码(Key)”</td>
      </tr>
      <tr>
    <td  colspan="2"><label>
      <input type="submit" name="Submit" value="保存" class="bnt" />
    </label></td>
  </tr>
</table>
</form>

</body>
</html>
<%
rs.nothing
set rs=nothing
 conn.Close
    Set conn = Nothing
%>

