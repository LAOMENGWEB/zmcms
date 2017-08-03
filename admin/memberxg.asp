<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="Inc/Function.asp"-->
<!--#include file="class_page.asp"-->
<!--#include file="../Inc/md5.asp"-->
<%
if Instr(session("AdminPurview"),"|23,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</script>"
  response.end
end if
'========判断是否具有管理权限
id=request.QueryString("id")
%>
<%
If request.Form("submit") = "修改" Then
Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from Members where ID="&ID
rs.open sql,conn,1,3
rs("MemName")=Request.Form("MemName")
rs("RealName")=trim(Request.Form("RealName"))
if trim(Request.Form("Password"))<>"" then
if len(trim(Request.Form("Password")))<6 or len(trim(Request.Form("Password")))>16  then
response.write "<script language=javascript> alert('会员密码必填，且字符数为6-16位！');history.back(-1);</script>"
response.end
end if
if Request.Form("Password")<>Request.Form("vPassword") then 
response.write "<script language=javascript> alert('两次输入的密码不一样！');history.back(-1);</script>"
response.end
end if
rs("Password")=Md5(Request.Form("Password"))
end if
rs("Sex")=Request.Form("Sex")
mGroupIdName=split(Request.Form("GroupID"),"┎╂┚")
rs("GroupID")=mGroupIdName(0)
rs("GroupName")=mGroupIdName(1)
rs("Company")=trim(Request.Form("Company"))
rs("Address")=trim(Request.Form("Address"))
rs("ZipCode")=trim(Request.Form("ZipCode"))
rs("Fax")=trim(Request.Form("Fax"))
rs("Mobile")=trim(Request.Form("Mobile"))
rs("Email")=trim(Request.Form("Email"))
rs("HomePage")=trim(Request.Form("HomePage"))
rs("lang")=lang
if Request.Form("Working")="1" then
rs("Working")=1
else
rs("Working")=0
end if
rs.update
rs.close
set rs=nothing 
call addlog("修改会员资料")
response.Write "<script language=javascript>alert(""恭喜您会员资料修改成功"");window.location.href=""member.asp""</script>"
end if
      set rs = server.createobject("adodb.recordset")
      sql="select * from Members where ID="& ID&" and lang='"&lang&"'"
      rs.open sql,conn,1,1
      If rs.EOF Then
    response.Write "<script language=javascript>window.location.href=""member.asp""</script>"
    response.End
    else
	  mMemName=rs("MemName")
	  mRealName=rs("RealName")
	  mSex=rs("Sex")
	  mGroupID=rs("GroupID")
	  mGroupName=rs("GroupName")
	  mCompany=rs("Company")
	  mAddress=rs("Address")
	  mZipCode=rs("ZipCode")
	  mFax=rs("Fax")
	  mMobile=rs("Mobile")
	  mEmail=rs("Email")
	  mHomepage=rs("Homepage")
	  mWorking=rs("Working")
  End if
	  rs.close
      set rs=nothing 
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>zmcms_管理页面</title>
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


  <form name="editMemForm" method="post" >


<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">会员信息修改  <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>


<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="huiyuan.asp" target="main">返回菜单</a></div></td>
 

 </tr>
</table>
 <table width="99%" border="0" align="center" cellpadding="0" cellspacing="0" class="dtable4">

      <tr>
        <td width="12%" >登&nbsp;录&nbsp;名：</td>
		<td width="88%">
        <input name="MemName" type="text"  id="MemName" style="WIDTH: 120;" value="<%=mMemName%>" maxlength="16" readonly />&nbsp;&nbsp;3-16位字符，不可修改</td>
      </tr>
      <tr>
        <td >真实姓名：</td>
		<td>
        <input name="RealName" type="text"  id="RealName" style="WIDTH: 120;" value="<%=mRealName%>" maxlength="16"></td>
      </tr>
      <tr>
        <td >密　　码：</td>
		<td>
        <input name="Password" type="password"  id="Password" maxlength="20" style="WIDTH: 120;">          &nbsp;&nbsp;6-16位字符，不填表未修改密码</td>
      </tr>
      <tr>
        <td >确认密码：</td>
		<td>
        <input name="vPassword" type="password"  id="vPassword" maxlength="20" style="WIDTH: 120;">          &nbsp;</td>
      </tr>
      <tr>
        <td >性　　别：</td>
		<td>
          <input type="radio" class="none" name="sex" value="先生" <%if mSex="先生" then response.write ("checked")%>>          &nbsp;先生&nbsp;
        <input type="radio" class="none" name="sex" value="女士" <%if mSex="女士" then response.write ("checked")%>>          &nbsp;女士</td>
      </tr>
      <tr>
        <td >会员组别：</td>
		<td>
		  <select name="GroupID" >
		  <% call SelectGroup() %>
          </select></td>
      </tr>
      <tr>
        <td >单位名称：</td>
		<td>
        <input name="Company" type="text"  id="Company" style="WIDTH: 240;" value="<%=mCompany%>" maxlength="100"></td>
      </tr>
      <tr>
        <td >地　　址：</td>
		<td>
        <input name="Address" type="text"  id="Address" style="WIDTH: 240;" value="<%=mAddress%>" maxlength="100"></td>
      </tr>
      <tr>
        <td >邮　　编：</td>
		<td>
        <input name="ZipCode" type="text"  id="ZipCode" style="WIDTH: 120;" value="<%=mZipCode%>" maxlength="16"></td>
      </tr>
      
      <tr>
        <td >传　　真：</td>
		<td>
        <input name="Fax" type="text"  id="Fax" style="WIDTH: 120;" value="<%=mFax%>" maxlength="16"></td>
      </tr>
      <tr>
        <td >移动电话：</td>
		<td>
        <input name="Mobile" type="text"  id="Mobile" style="WIDTH: 120;" value="<%=mMobile%>" maxlength="16"></td>
      </tr>
      <tr>
        <td >电子邮箱：</td>
		<td>
        <input name="Email" type="text"  id="Email" style="WIDTH: 240;" value="<%=mEmail%>" maxlength="50"></td>
      </tr>
      <tr>
        <td >网　　址：</td>
		<td>
        <input name="HomePage" type="text"  id="HomePage" style="WIDTH: 240;" value="<%=mHomePage%>" maxlength="50"></td>
      </tr>
      <tr>
        <td >生　　效：</td>
		<td>
        <input class="none" name="Working" type="checkbox"  value="1"  <%if mWorking=1 then response.write ("checked")%>></td>
      </tr>

      <tr>
        <td colspan="2"><input type="submit" name="Submit" class="bnt"  id="submitSaveEdit" value="修改"  ></td>
      </tr>
    </table>
  </form>

</body>
</html>

<% 
sub SelectGroup()
  dim rs,sql
  set rs = server.createobject("adodb.recordset")
  sql="select GroupID,GroupName from hyGroup where lang='"&lang&"'"
  rs.open sql,conn,1,1
  if rs.bof and rs.eof then
    response.write("未设组别")
  end if
  while not rs.eof
    response.write("<option value='"&rs("GroupID")&"┎╂┚"&rs("GroupName")&"'")
    if mGroupID=rs("GroupID") then response.write ("selected")
    response.write(">"&rs("GroupName")&"</option>")
    rs.movenext
  wend
  rs.close
  set rs=nothing
  session("mGroupID")=zmcmsGroupID
end sub
%>