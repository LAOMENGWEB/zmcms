<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="Inc/Function.asp"-->
<%

if Instr(session("AdminPurview"),"|1091,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</script>"
response.end
end if


del=request("del")
if del<>"" then
Set fso = Server.CreateObject("Scripting.FileSystemObject")
fso.DeleteFolder server.MapPath("../"&del)
set fso=nothing
call addlog("修改后台目录")
response.write "<script language='javascript'>"
response.write "alert('修改成功，请牢记新的后台目录名称，下次登陆时请以新的后台目录名称登陆！');"
response.write "location.href='login.asp';"
response.write "</script>"
response.end
end if
Set fso = Server.CreateObject("Scripting.FileSystemObject")
set fd  = fso.getfolder(server.MapPath("./"))
fdname=fd.name
set fso=nothing
action=request("ren")
if action="" then 
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><head>
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
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">后台目录修改</div></td>
</tr>
</table>


<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="anquan.asp" target="main">返回菜单</a></div></td>
 

 </tr>
</table>
<table  width="99%" align="center"  style=" background:#f9f9f9">
  <tr > 
    <td colspan="2"><p>·此操作需FSO功能支持。</p>
      <p> ·执行此操作前，请先关闭其它网页窗口，断开FTP连接。      </p>
      <p>·提交后，请耐心等待（请不要重复点击上面的按钮），完成时间视网络状况而定。</p>
      <p> ·请在网络空闲时进行此操作，执行此操作可能导致服务器变慢或不稳定    </p>
    <p>·请勿包含&*%/:?"|\等特殊符号 </p></td>
  </tr>
</table>

<table width="99%" border="0" align="center" cellpadding="0" cellspacing="0" class="dtable4">
 <form action="" method=post name="adminpath">
  <tr>
    <td colspan="2" style="padding:10px">
 
      <strong>修改后台目录:</strong>    
      <input type=text  class=i_text  value="<%=fdname%>" name="newpath" size=30   maxlength="250">    
   </Td>
    </TR>
	 <tr>
    <td colspan="2" style="padding:10px">
	  
          <INPUT name="ren" TYPE="hidden" value="ok">
          <INPUT class=bnt type=submit value=确定 name=eventSubmit_doSavemailcfg  > 
       
		</Td>
    </TR>
		</form>
</TABLE>





    



<%
elseif action="ok" then 
newpath=trim(Request.Form("newpath"))
if newpath = "" then
response.write "<script language='javascript'>"
response.write "alert('您什么也没有填写，请检查后重新提交！');"
response.write "location.href='javascript:history.go(-1)';"
response.write "</script>"
response.end
elseif zmcmscheck(request("newpath"))<>request("newpath") then
response.write "<script language='javascript'>"
response.write "alert('您填写的内容中含有非法字符，请检查后重新输入！');"
response.write "location.href='javascript:history.go(-1)';"
response.write "</script>"
response.end
elseif newpath=fdname then
response.write "<script language='javascript'>"
response.write "alert('新旧名称相同，未作修改！');"
response.write "location.href='javascript:history.go(-1)';"
response.write "</script>"
response.end
else
Set fso = Server.CreateObject("Scripting.FileSystemObject")
if fso.FolderExists(server.MapPath("../"&newpath))=true then
response.write "<script language='javascript'>"
response.write "alert('这个目录已经存在，请输入其它名称。');"
response.write "location.href='javascript:history.go(-1)';"
response.write "</script>"
response.end
else
fso.CopyFolder server.MapPath("./"),server.MapPath("../"&newpath)
newpath="../"&newpath&"/managexg.asp?del="&fdname
Response.Redirect(newpath)  '通过地址栏递交变量，转向到新的目录，并删除老目录
end if
set fso=nothing
end if
end if
%></td>
  </tr>

</table>

<br />
</body>
</html>
<%function zmcmscheck(txt)
zmcmscheck=txt
chrtxt="33|34|35|36|37|38|39|40|41|42|43|44|47|58|59|60|61|62|63|91|92|93|94|96|123|124|125|126|128"
chrtext=split(chrtxt,"|")
for c=0 to ubound(chrtext)
zmcmscheck=replace(zmcmscheck,chr(chrtext(c)),"")
next
end function%>