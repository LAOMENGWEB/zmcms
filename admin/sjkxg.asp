<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<%
if Instr(session("AdminPurview"),"|1091,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</script>"
  response.end
end if
'========判断是否具有管理权限
%>
<%
action=request("ren")
if action="" then
%>
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
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">数据库名称修改</div></td>
</tr>
</table>


<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="anquan.asp" target="main">返回菜单</a></div></td>
 

 </tr>
</table>
<table  width="99%" align="center" style=" background:#f9f9f9">
  <tr > 
    <td colspan="2" style="border-bottom:none"><p>·此操作需FSO功能支持。</p>
      <p> ·执行此操作前，请先关闭其它网页窗口，断开FTP连接。      </p>
      <p>·提交后，请耐心等待（请不要重复点击上面的按钮），完成时间视网络状况而定。</p>
      <p> ·请在网络空闲时进行此操作，执行此操作可能导致服务器变慢或不稳定    </p>
    <p>请勿包含&amp;*%/:?&quot;|\等特殊符号, </p>
	<p>修改的时候请勿修改相同的名称，否则有可能出错，例如，zmcms.asp 和zmcms 这也是相同的名称。 </p>
	
	</td>
  </tr>
</table>

<form action=sjkxg.asp method=post name=backup>
  
<TABLE width="99%" border=0 align="center" cellPadding=0 cellSpacing=0  class="dtable4">
  <TBODY>
  

  <TR>
    <TD><div align="left">数据库名字修改:
      <input type=hidden name="oldname" value="<%=replace(db,left(db,InStrRev(db,"/")),"")%>">
        <input  name="newnamem" type=text  class=i_text value="<%=dataname%>" size=30   maxlength="250">    
    </div></TD>
    </TR>
	  <TR>
    <TD>
  
        <div align="left">
          <input type="hidden" name="ren" value="ok">
          <input title='对网站正在使用的数据库文件更名' type="submit" class="bnt" name="Submit" value="在线更名"  style="cursor:hand" onClick="{if(confirm('为数据库文件改个文件名，可防止数据库文件被恶意下载。\n\n 此操作实际上是系统分两步自动完成的（先复制再删除）。\n\n如果在线更名过程中出错，请通过原始方法修改（即FTP登陆后修改）。\n\n单击确定继续，单击取消返回。')){this.document.backup.submit();return true;}return false;}"> 
        </div></D>
    </TR>
</TBODY>
</TABLE>







</form>
<% 
elseif action="ok" then
newnamem=trim(request("newnamem"))
oldname=trim(request("oldname"))
if newnamem="" then
response.write "<script language='javascript'>"
response.write "alert('出错了，未填写新的数据库文件名。');"
response.write "location.href='javascript:history.go(-1)';"
response.write "</script>"
response.end
end if
if oldname=newnamem then
response.write "<script language='javascript'>"
response.write "alert('新旧文件名相同，未作修改！');"
response.write "location.href='javascript:history.go(-1)';"
response.write "</script>"
response.end
end if

if instr(lcase(newnamem)," ")>0 or instr(lcase(newnamem),"o.f")>0 or instr(lcase(newnamem),"e.b")>0 or instr(lcase(newnamem),"n.t")>0 or instr(lcase(newnamem),"r.m")>0 or instr(lcase(newnamem),"o.asp")>0 or instr(lcase(newnamem),"t.asp")>0 or instr(lcase(newnamem),"s.asp")>0 then
response.write "<script language='javascript'>"
response.write "alert('出错了，新的文件名中包含与系统相冲突的字符，请使用其它名称。');"
response.write "location.href='javascript:history.go(-1)';"
response.write "</script>"
response.end
end if
if zmcmscheck(request("newnamem"))<>request("newnamem") then
response.write "<script language='javascript'>"
response.write "alert('出错了，文件名中含有非法字符，请检查后重新输入！');"
response.write "location.href='javascript:history.go(-1)';"
response.write "</script>"
response.end
end if
oldpath=server.mappath(db)
newpath=replace(oldpath,oldname,newnamem)
Set fso = CreateObject("Scripting.FileSystemObject")
if fso.fileexists(oldpath) then
fso.CopyFile oldpath,newpath
else
response.write "<script language='javascript'>"
response.write "alert('操作失败，请刷新后重试！');"
response.write "location.href='javascript:history.go(-1)';"
response.write "</script>"
response.end
end if
set fso=nothing
Set fso = CreateObject("Scripting.FileSystemObject")
set connold=Fso.OpenTextFile(server.mappath("../inc/conn.asp"))
conntxt=replace(connold.readall,oldname,newnamem)
Connold.close
set connnew=Fso.OpenTextFile(server.mappath("../inc/conn.asp"),2)
connnew.write(conntxt)
Connnew.close
set fso=nothing
response.write "<script language='javascript'>"
response.write "location.href='sjkxg.asp?del="&oldname&"';"
response.write "</script>"
response.end
end if
if request("del")<>"" then
oldname=request("del")
newnamem=replace(db,left(db,InStrRev(db,"/")),"")
newpath=server.mappath(db)
oldpath=replace(newpath,newnamem,oldname)
Set fso = CreateObject("Scripting.FileSystemObject")
if fso.fileexists(oldpath) then
fso.DeleteFile(oldpath)
call addlog("修改数据库名称")
response.write "<script language='javascript'>"
response.write "alert('操作成功，数据库文件改名成功！');"
response.write "location.href='sjkxg.asp';"
response.write "</script>"
response.end
else
response.write "<script language='javascript'>"
response.write "alert('操作失败，请刷新后重试！');"
response.write "location.href='javascript:history.go(-1)';"
response.write "</script>"
response.end
end if
set fso=nothing

end if
%>
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