<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="admin.asp"-->
<!--#include file="conn.asp"-->


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>zmcms_管理页面</title>
 <link href="css/style.css" rel="stylesheet" type="text/css">
	 <style type="text/css">
.table1{ font-size: 14px; background: #F9f9f9;border:none;}
.table1 td{padding:10px; font-family:"微软雅黑";color:#ccc; font-weight:bold;}
.table1 td a{color:#ffffff}
</style>
</head>
<body>


<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center"> <%if request.QueryString("action")="config"   then%> SQL注入设置<%end if%>
 <%if request.QueryString("action")= ""  then%> SQL注入信息查看<%end if%></div></td>
</tr>
</table>


<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="anquan.asp" target="main">返回菜单</a></div></td>
 

 </tr>
</table>
<%
if Instr(session("AdminPurview"),"|109,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</scrip>"
  response.end
end if
'========判断是否具有管理权限
%>
<%
Server.ScriptTimeout	=500						
URL= Request.ServerVariables("URL")
Action= Request("Action")



	Select Case Action
		Case "Del"
			Call Delip()
			call addlog("删除SQL注入信息")
		Case "lock"
			Call lockIP()
			
		Case "unlock"
			Call UnLockip()
			
		Case "Logout"
			Call Logout()
		Case "config"
			Call config()
			call addlog("设置SQL注入")
		Case "saveconfig"
			Call saveconfig()
		Case Else 
			Call Main()
	end Select


Sub Login()
	%>
	 <%
End Sub

Sub Main()
	Call header()
%>
<script language="JavaScript">
<!--
function CheckAll(form)  {
  for (var i=0;i<form.elements.length;i++)    {
    var e = form.elements[i];
    if (e.name != 'chkall')       e.checked = form.chkall.checked; 
   }
  }
//-->
</script> 

<style type="text/css">
<!--
.STYLE1 {
	font-size: 14px;
	font-weight: bold;
}
-->
</style>


    <%
sql="select * from SqlIn order by id desc"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
if rs.eof and rs.bof then
%>
<table width="99%" align="center" class="table1" style="margin-top:10px">
  <tr>
    <td align="center" style="border-right:none;border-bottom:none">暂无安全记录</td>
  </tr>
</table>

<%
else
'分页的实现 
listnum=20
Rs.pagesize=listnum
page=Request("page")
if (page-Rs.pagecount) > 0 then
page=rs.pagecount
elseif page = "" or page < 1 then
page = 1
end if
rs.absolutepage=page
'编号的实现
j=rs.recordcount
j=j-(page-1)*listnum
i=0
nn=request("page")
if nn="" then
n=0
else
nn=nn-1
n=listnum*nn
end if%>

<table width="99%"  align="center" cellpadding="0" cellspacing="0" class="dtable2">
  <tr class="wz">
    <td width="5%"  >全选</td>

    <td width="4%"  >编号</td>
    <td width="16%"  >攻击ＩＰ</td>
    <td width="7%"  >当前状态</td>
    <td width="8%"  >是否锁定</td>
    <td width="14%"  >操作页面</td>
    <td width="9%"  >操作时间</td>
    <td width="7%"  >提交方式</td>
    <td width="8%"  >提交参数</td>
    <td width="22%"  >提交数据</td>
  </tr>
  <%do while not rs.eof and i<listnum
n=n+1%>
  <form action="?del=ok" method="post" name="check" id="check">
    <tr align="center" height="22">
      <td ><input name="ID" type="checkbox" id="ID" value="<%=rs("id")%>" class="none"/></td>
      <td >
          <%=rs("id")%></td>
      <td ><%=rs("SqlIn_IP")%></td>
      <td ><%	if rs("Kill_ip")=true then 
			response.write "<font color='red'>已锁定</font>"
		else
			response.write "<font color='green'>已解锁</font>"
		end if
	%></td>
      <td ><%	if rs("Kill_ip")=true then 
			response.write "<a href="&URL&"?action=unlock&id="&rs("id")&" color:#FF0000"">解锁IP</a>"
		else
			response.write "<a href="&URL&"?action=lock&id="&rs("id")&" color:#006600"">锁定IP</a>"
		end if
	%>      </td>
      <td ><%=rs("SqlIn_WEB")%></td>
      <td ><%=rs("SqlIn_TIME")%></td>
      <td ><%=rs("SqlIn_FS")%></td>
      <td ><%=rs("SqlIn_CS")%></td>
      <td style="border-right:none;"><%=N_Replace(rs("SqlIn_SJ"))%></td>
    </tr>
    <%rs.movenext 
i=i+1 
j=j-1
loop%>
    <tr>
      <%filename=URL%>
	  <td><input name="chkall" type="checkbox" id="chkall" value="select" onClick="CheckAll(this.form)" class="none"/></td>
      <td colspan="9" style="padding:5px; text-align:left"> <input type="submit" name="action" value="删除" class="bnt"> 
        共<%=Rs.recordcount%> 条记录&nbsp;&nbsp;<%=listnum%> 条记录/页&nbsp;&nbsp;共 <%=rs.pagecount%> 页
        <% if page=1 then %>
          <%else%>
          <a href="<%=filename%>"><strong>|&lt;&lt;</strong></a> <a href="<%=filename%>?page=<%=page-1%>"><strong>&lt;&lt;</strong></a> <a href="<%=filename%>?page=<%=page-1%>"><b>[<%=page-1%>]</b></a>
          <%end if%>
          <% if rs.pagecount=1 then%>
          <%else%>
          <b>[<%=page%>]</b>
          <%end if%>
          <% if rs.pagecount-page <> 0 then %>
          <a href="<%=filename%>?page=<%=page+1%>"><b>[<%=page+1%>]</b></a> <a href="<%=filename%>?page=<%=page+1%>"><strong>&gt;&gt;</strong></a> <a href="<%=filename%>?page=<%=rs.pagecount%>"><strong>&gt;&gt;|</strong></a>
          <%end if%>
        
      </td>
      <%end if%>
    </tr>
  </form>
</table>
<%
	Call footer()
end Sub

sub config()
	Call header()
	Set rsinfo=conn.execute("select * from sqlconfig")
	N_In		= rsinfo("N_In")
	Kill_IP		= rsinfo("Kill_IP")			
	WriteSql	= rsinfo("WriteSql")		
	alert_url	= rsinfo("alert_url")
	alert_info	= rsinfo("alert_info")
	kill_info	= rsinfo("kill_info")
	N_type		= rsinfo("N_type")
	Sec_Forms	= rsinfo("Sec_Forms")
	Sec_Form_open = rsinfo("Sec_Form_open")
	rsinfo.close
	Set rsinfo=Nothing 

%>

<table width="99%"   align="center" cellpadding="0" cellspacing="0"  class="dtable4">
            <form name="form" method="post" action="<%=url%>?action=saveconfig">
         <tr   >
                <td width="14%" >需要过滤的关键字：</td>
                <td width="86%">
                    <input name="N_In" type="text" value="<%=N_In%>" id="r_str"  size="50">
           用&quot;|&quot;分开</td>
              </tr>
              <tr >
                <td >是否记录入侵者信息：</td><td>
                    <select name="WriteSql" id="r_kill">
                      <option value="1" <%if WriteSql=1 Then response.write "selected"%>>是</option>
                      <option value="0" <%if WriteSql=0 Then response.write "selected"%>>否</option>
                  </select></td>
              </tr>
                <tr >
                <td>是否启用锁定IP：</td><td>
                    <select name="Kill_IP" id="r_kill">
                      <option value="1" <%if Kill_IP=1 Then response.write "selected"%>>是</option>
                      <option value="0" <%if Kill_IP=0 Then response.write "selected"%>>否</option>
                    </select></td>
              </tr>
              <tr >
                <td>是否启用安全页面：</td><td>
                    <select name="Sec_Form_open" id="r_kill">
                      <option value="1" <%if Sec_Form_open=1 Then response.write "selected"%>>是</option>
                      <option value="0" <%if Sec_Form_open=0 Then response.write "selected"%>>否</option>
                    </select>
                慎用这个功能，除非你对确认此页面无需过滤，并确定对安全没影响！ </td>
              </tr>
                <tr   >
                <td>您认为安全的页面：</td><td>
                    <input name="Sec_Forms" type="text" value="<%=Sec_Forms%>" id="r_str"  size="50">
                  用&quot;|&quot;分开</td>
              </tr>
              <tr >
                <td >出错后的处理方式：</td><td>
                    <select name="N_type" id="r_kill">
                      <option value="1" <%if N_type=1 Then response.write "selected"%>>直接关闭网页</option>
                      <option value="2" <%if N_type=2 Then response.write "selected"%>>警告后关闭</option>
                      <option value="3" <%if N_type=3 Then response.write "selected"%>>跳转到指定页面</option>
                      <option value="4" <%if N_type=4 Then response.write "selected"%>>警告后跳转</option>
                  </select></td>
              </tr>
                 <tr   >
                <td>出错后跳转Url：</td><td>
                   <input name="alert_url" type="text" value="<%=alert_url%>" id="r_str"  size="30"></td>
              </tr>
              <tr >
                <td >警告提示信息：</td><td>
                  <textarea name="alert_info" cols="45" rows="4" id="alert_info" style=";  "><%=alert_info%></textarea>                  
                \n\n换行</td>
              </tr>
                 <tr   >
                <td>阻止访问提示信息：</td><td>
                    <textarea name="kill_info" cols="45" rows="4" id="r_err2" style=";  "><%=kill_info%></textarea>
                  \n\n换行 </td>
              </tr>
	
              <tr >
                <td colspan="2"><input name="enter_3" type="submit" id="enter_3" class="bnt" value="保存设置"  ></td>
              </tr>
            </form>
</table>
<%
	Call footer()
end Sub

Sub header()

%>




<%

End Sub 
sub footer()

end Sub
Sub Delip()
dim id 
conn.execute("delete from SqlIn where id in ( " & id & ")")
Response.Redirect URL
End sub

Sub Lockip()
id = clng(request("id"))
if sqlno=1 then
conn.execute("update SqlIn set Kill_ip=true where id="&id)
else
conn.execute("update SqlIn set Kill_ip=1 where id="&id)
end if
call addlog("锁定注入IP")
Response.Redirect URL
End sub

Sub UnLockip()
id = clng(request("id"))
if sqlno=1 then
conn.execute("update SqlIn set Kill_ip=False where id="&id)
else
conn.execute("update SqlIn set Kill_ip=0 where id="&id)
end if
call addlog("解开注入IP")
Response.Redirect URL
End sub

Sub Logout()
	Session("AdminPassWord")="NUll"
	Response.Redirect URL
End Sub

Sub saveconfig
	N_In		=replace(request.form("N_In"),"'","''")
	Kill_IP		=request.form("Kill_IP")			
	WriteSql	=request.form("WriteSql")		
	alert_url	=request.form("alert_url")
	alert_info	=request.form("alert_info")
	kill_info	=request.form("kill_info")
	N_type		=request.form("N_type")
	Sec_Forms	=request.form("Sec_Forms")
	Sec_Form_open=request.form("Sec_Form_open")

	sql="update sqlconfig set N_In='"&N_In&"',Kill_IP="&Kill_IP&",WriteSql="&WriteSql&",alert_url='"&alert_url&"',alert_info='"&alert_info&"',kill_info='"&kill_info&"',N_type="&N_type&",Sec_Forms='"&Sec_Forms&"',Sec_Form_open="&Sec_Form_open&""
		Response.Write "<script language='javascript'>alert('保存成功！');document.location.href('admin_sql.asp?Action=config');</script>"
	conn.execute(sql)
	Application.Lock
	set Application("Neeao_config_info")=nothing
	Application.unlock
	Call main()
End Sub 

Function N_Replace(N_urlString)
	N_urlString = Replace(N_urlString,"'","''")
    N_urlString = Replace(N_urlString, ">", "&gt;")
    N_urlString = Replace(N_urlString, "<", "&lt;")
    N_Replace = N_urlString
End Function
%>
<%
if Request.QueryString("del")="ok" then
if Request("id")="" then
Response.Write "<script>alert('请选择要删除的记录!');window.location.href='admin_sql.asp';</script>" 
response.end()
end if
dim sql
sql="delete from [SqlIn] where id in ("&Request("id")&")"
conn.Execute ( sql )
conn.close
set conn=nothing
Response.Write "<script>alert('批量删除成功!');window.location.href='admin_sql.asp';</script>" 
end if
%>

</body>
</html>