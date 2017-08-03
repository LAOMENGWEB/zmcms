<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!-- #include file="admin.asp" -->
<!-- #include file="Inc/function.asp" -->

<%
if Instr(session("AdminPurview"),"|301,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</script>"
  response.end
end if
'========判断是否具有管理权



%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>zmcms</title>
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
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center"> <%if request.QueryString("action")="BackupData" then%> 数据库备份 <%end if%>
 <%if request.QueryString("action")="CompressData" then%> 数据库压缩 <%end if%>
 <%if request.QueryString("action")="ShowRestore" then%> 数据库恢复 <%end if%></div></td>
</tr>
</table>


<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="anquan.asp" target="main">返回菜单</a></div></td>
 

 </tr>
</table>
<%
select case request("action")
 
	case "CompressData"
	
			CompressData()
			call addlog("压缩数据库")
			
		case "ShowRestore"
	
			call ShowRestore()
			call addlog("数据库删除")	
			
		Case "delback"
	
	call DelBackup()
	case "BackupData"
		
	    if request("act")="Backup" then
		
				call updata()
		
		else
			call addlog("备份数据库")
				call BackupData()
			
		end if
	case "RestoreData"
		
	
			
				dim Dbpath,backpath
				Dbpath="../Databackup/"&Request.QueryString("file")
				backpath=request.form("backpath")
				if dbpath="" then
				Call Alert("请输入您要恢复的数据库路径及名称!","-1")	
				else
				Dbpath=server.mappath(Dbpath)
				end if
				backpath=server.mappath(".."&db)
			
				Set Fso=Server.CreateObject("Scripting.FileSystemObject")
				if fso.fileexists(dbpath) then  					
					fso.copyfile Dbpath,Backpath
					Call Alert("成功恢复数据!","-1")	
					call addlog("恢复数据库")
				else
					Call Alert("备份目录下并无您的备份文件!","-1")	
				end if
			
	
end select
%>
<%
'====================压缩数据库 =========================
sub CompressData()
%>
<table width="99%" border="0" align="center" cellpadding="0"  cellspacing="0" class="dtable4">
<form action="datebf.asp?action=CompressData" method="post">
<tr>
<td  style="border-right:none;" ><b>注意：</b>输入数据库所在相对路径,并且输入数据库名称（正在使用中数据库可能会压缩失败，请选择备份数据库进行压缩操作） </td>
</tr>
<tr>
<td   style="border-right:none;">压缩数据库：
  <input name="dbpath" type="text" value="<%=Db%>" size="50">
&nbsp;
<input type="submit" class="bnt" value="开始压缩"></td>
</tr>
<tr>
<td  style="border-right:none" ><input name="boolIs97" type="checkbox" class="noborder" value="True" style="border:none">如果使用 Access 97 数据库请选择
(默认为 Access 2000 数据库)<br></td>
</tr>
<tr>
  <td >注：请尽量用ftp下载回数据库后压缩，以免出错！如果你非要使用此功能，请备份后再压缩！<strong>数据库出错本程序作者概不负责!</strong></td>
</tr>
</form>
</table>
<%
dim dbpath,boolIs97
dbpath = request("dbpath")
boolIs97 = request("boolIs97")

If dbpath <> "" Then
dbpath = server.mappath(dbpath)
	response.write(CompactDB(dbpath,boolIs97))
End If

end sub

'=====================压缩参数=========================
Function CompactDB(dbPath, boolIs97)
	Dim fso, Engine, strDBPath,JET_3X
	strDBPath = left(dbPath,instrrev(DBPath,"\"))
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	
	If fso.FileExists(dbPath) Then
		fso.CopyFile dbpath,strDBPath & "temp.mdb"
		Set Engine = CreateObject("JRO.JetEngine")
	
		If boolIs97 = "True" Then
			Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp.mdb", _
			"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp1.mdb;" _
			& "Jet OLEDB:Engine Type=" & JET_3X
		Else
			Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp.mdb", _
			"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp1.mdb"
		End If
	
		fso.CopyFile strDBPath & "temp1.mdb",dbpath
		fso.DeleteFile(strDBPath & "temp.mdb")
		fso.DeleteFile(strDBPath & "temp1.mdb")
		Set fso = nothing
		Set Engine = nothing
		
		Call Alert("压缩成功!","-1")
	Else
		Call Alert("数据库名称或路径不正确. 请重试!","-1")
	End If
End Function
%>
<%
sub ShowRestore()%>
<TABLE width="99%" border=0 align=center cellpadding="0" cellSpacing=0  class="dtable2">
  <TBODY>
    <TR class="wz">
      <TD width="42%" >备份文件名</TD>
      <TD width="43%" >备份时间</TD>
      <TD width="15%" >操作</TD>
    </TR>
    <% Dim Fso,i
	Set Fso=server.createobject("Scripting.FileSystemObject")
	dim theFolder,theFile,strFileType
	if Not fso.FolderExists(Server.MapPath("../Databackup")) then
		response.write "找不到数据库备份文件夹"
		response.end
	end if
	Set theFolder=fso.GetFolder(Server.MapPath("../Databackup"))

	Set Files = theFolder.Files 
	If Files.Count=0  then
	%>
	
	  <TR>
      <TD colspan="3" align="center" >	<%
		response.write "找不到数据备份"
		response.end
		%></TD>
    </TR>
<%
	end if
	For Each theFile In theFolder.Files%>
    <TR >
      <TD align="center" ><%=theFile.Name%></TD>
      <TD align="center" ><%=theFile.DateLastModified%></TD>
      <TD align="center" style="border-right:none;"><a href="datebf.asp?action=RestoreData&file=<%=theFile.Name%>">还原</a> | <a href="datebf.asp?action=delback&file=<%=theFile.Name%>">删除</a></TD>
	</TR>
	<%Next
Set Fso=Nothing%>
  </TBODY>
</TABLE>
<%end sub%>




<%
'====================备份数据库=========================
sub BackupData()
%>
	<table width="99%" border="0" align="center" cellpadding="0"  cellspacing="0"  class="dtable4">
	  <form method="post" action="datebf.asp?action=BackupData&act=Backup">
		<tr>
		  <td width="19%"  align="center"  >当前数据库路径(相对)：</td>
		  <td width="81%"  ><input name=DBpath type=text id="DBpath" value="<%=Db%>" size="40"  Readonly="true"/></td>
	    </tr>
		<tr>
		  <td  align="center"  >备份数据库目录(相对)：</td>
		  <td  style="border-right:none"><input name=bkfolder type=text value="../Databackup/" size="40" Readonly="true"/>
&nbsp;如目录不存在，程序将自动创建</td>
	    </tr>
		<tr>
		  <td  align="center"  >备份数据库名称(名称)：</td>
		  <td  style="border-right:none"><input name=bkDBname type=text value="<%=year(Now)&month(Now)&day(Now)&"_"&hour(Now)&Minute(Now)&Second(Now)%>.mdb" size="40"  Readonly="true"/>
&nbsp;如备份目录有该文件，将覆盖，如没有，将自动创建</td>
	    </tr>
		<tr>
	
		  <td   colspan="2"><input type=submit class="bnt" value="确定" /></td>
	    </tr>
		<tr>
		  
		  <td colspan="2">注意：所有路径都是相对与程序空间根目录的相对路径 (需要FSO支持，FSO相关帮助请看微软网站)</td>
	    </tr>	
		</form>
</table>
<%
end sub

sub updata()
	dim Dbpath,bkfolder,bkdbname,fso
	Dbpath=""&db&""
	Dbpath=server.mappath(Dbpath)
	bkfolder=request.form("bkfolder")
	bkdbname=request.form("bkdbname")
	Set Fso=Server.CreateObject("Scripting.FileSystemObject")
	if fso.fileexists(dbpath) then
		If CheckDir(bkfolder) = True Then
		fso.copyfile dbpath,bkfolder& "\"& bkdbname
		else
		MakeNewsDir bkfolder
		fso.copyfile dbpath,bkfolder& "\"& bkdbname
		end if		
		Call Alert ("备份数据库成功!","-1")
	Else
		Call Alert ("数据库路径错误","-1")
	End if
end sub
'------------------检查某一目录是否存在-------------------
Function CheckDir(FolderPath)
    dim fso1
	folderpath=Server.MapPath(".")&"\"&folderpath
    Set fso1 = Server.CreateObject("Scripting.FileSystemObject")
    If fso1.FolderExists(FolderPath) then
       '存在
       CheckDir = True
    Else
       '不存在
       CheckDir = False
    End if
    Set fso1 = nothing
End Function
'-------------根据指定名称生成目录-----------------------
Function MakeNewsDir(foldername)
	dim f,fso1
    Set fso1 = Server.CreateObject("Scripting.FileSystemObject")
        Set f = fso1.CreateFolder(foldername)
        MakeNewsDir = True
    Set fso1 = nothing
End Function
%>
<%
'====================恢复数据库=========================
sub RestoreData()
%>
<table width="99%" border="0" align="center" cellpadding="0"  cellspacing="0" >

<form method="post" action="datebf.asp?action=RestoreData&act=Restore">
<tr>
  <td width="19%"  align="center"  >备份数据库路径(相对)：</td>
  <td width="81%"  ><select name="DBpath" id="DBpath">
  <%
  Dim bflist
  bflist=FileList("../Databackup","mdb")
  If bflist<>"" then
  bflist1=split(left(bflist,len(bflist)-1),"|")
  For i=0 to Ubound(bflist1)
  %>
    <option value="../DataBackup/<%=bflist1(i)%>"><%=bflist1(i)%></option><%
	Next
	Else
	%><option value="">很遗憾！没有备份数据库</option><%
	End if
	%>
  </select>
  </td>
</tr>
<tr>
  <td  align="center"  >目标数据库路径(相对)：</td>
  <td  ><input type=text size=40 name=backpath value="<%=Db%>"  Readonly="true"/></td>
</tr>
<tr>
  <td   >&nbsp;</td>
  <td  ><input type=submit class="bnt" value="恢复数据" /></td>
</tr>
<tr>
  <td   >&nbsp;</td>
  <td  >请谨慎操作！出现问题概不负责！</td>
</tr>	
</form>
</table>
<%
end sub

%>

<%






Sub DelBackup()
	dim Document,fso
	backpath="../Databackup/"&Request.QueryString("file")
	
	Set Fso=server.createobject("Scripting.FileSystemObject")
	fso.DeleteFile(server.mappath(backpath))
	Set Fso=Nothing
	
 response.Redirect("datebf.asp?action=ShowRestore")
End sub




Function Alert(message,gourl) 
    message = replace(message,"'","\'")
    If gourl="-1" then
        Response.Write ("<script language=javascript>alert('" & message & "');history.go(-1)</script>")
    ElseIf gourl="-2" then
        Response.Write ("<script language=javascript>alert('" & message & "');history.go(-2)</script>")
    ElseIf gourl="Close" then
		Response.Write ("<script language=javascript>alert('" & message & "');window.opener=null;window.close();</script>")
	Else
        Response.Write ("<script language=javascript>alert('" & message & "');location='" & gourl &"'</script>")
    End If
    Response.End()
End Function
%>



</body>
</html>