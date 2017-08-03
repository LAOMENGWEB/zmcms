<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<%
if request("template")="default" Then
set rs=server.createobject("adodb.recordset")
sql="select * from wzset where lang='"&lang&"' ORDER BY id DESC"
rs.open sql,conn,1,3
rs("skins")=request("choose")
rs.update
rs.close
set rs=nothing
response.write "<script language=javascript> alert('设置成功！');location.replace('template.asp');</script>"
End If
Response.Buffer=true 
On Error Resume Next
Server.ScriptTimeOut = 5000
Response.Expires=0 
Dim StartTime,IsReplace,ImageFolder
IsReplace = true
ImageFolder = "fsoimg"
StatrTime = Timer()
'****************************************
'函数定义部分结束
'****************************************
Dim Fso,FsoFile,FileType,FileSize,FileTime,Path
Dim Dir
action=Trim(Request.QueryString("action"))
Set Fso=Server.CreateObject("Scripting.FileSystemObject")
IsErr
If action = "Del" then
Call DelAll
ElseIf action = "NewFile" then
Call NewFile
ElseIf action = "NewFolder" then
Call NewFolder
ElseIf action = "Rname" then
Call Rname
ElseIf action = "Edit" then
Call Edit
ElseIf action = "Save" then
Call Edit
Else
Dir=Trim(Request.QueryString("Dir"))
Path = Server.MapPath(""&sitepath&"template/") & Dir
Set FsoFile = Fso.GetFolder(Server.MapPath(""&sitepath&"template/"))
FsoFileSize = FsoFile.size	'空间大小统计
Set	FsoFile = nothing
Set	FsoFile = Fso.GetFolder(Path)
%>
<html>
<head>
<title>列遍模板目录</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
	<script language="JavaScript" src="js/jquery-1.7.2.min.js"></script>
<link href="css/style.css" rel="stylesheet" type="text/css">
	<style type="text/css">
.table1{ font-size: 14px; background: #F9f9f9;border:none;}
.table1 td{padding:10px; font-family:"微软雅黑";color:#ccc; font-weight:bold;}
.table1 td a{color:#ffffff}
</style>
<!--#include file="hqyy.asp"-->
</head>
<body >
<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">模板管理 <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>
<div style="background-color: #f9f9f9;width:100%; margin-top: 10px; float:left">
<div class="moban">
<ul>
<%
	For Each DirFolder in FsoFile.SubFolders
	FolderName=DirFolder.name
	FolderSize=GetFileSize(DirFolder.size)
	FolderTime=DirFolder.DateLastModified

	set rs = server.createobject("adodb.recordset")
	sql="select * from wzset where lang='"&lang&"' ORDER BY id DESC"
	rs.open sql,conn,1,1

	do while not rs.eof
        mobanurl = rs("skins")
        if mobanurl = FolderName then
		mobanshiyong = "<font color='green'>√ 当前在使用</font>,<a href='mb.asp?t=1'>编辑该模板</a>"
			else
		mobanshiyong = "<a href='template.asp?template=default&choose="&FolderName&"'>使用</a>"
	end if
	rs.MoveNext 
	loop

                   fpath=server.MapPath ("../template/"&FolderName&"/index.png")
                   if fso.FileExists(fpath) then
		    		mobanurl2 = "../template/"&FolderName&"/index.png"
                     else
		    		mobanurl2 = ""&sitepath&"images/template.png"
                   end if
%>
<li>
<div class="moban_left"><img src="<%=mobanurl2%>" widht="180" height="180"></div>
<div class="moban_right">
	<p>位置:../template/<%=FolderName%>/</p>
<%				
Set fs = Server.CreateObject("Scripting.FileSystemObject")
File = Server.MapPath("../template/"&FolderName&"/Author.txt")
Set txt = fs.OpenTextFile(File) 
If Not txt.atEndOfStream Then
Content = txt.ReadAll
Lines = Replace(Content, vbCrlf, "<br>")
Response.Write Lines
End If
%>

				<p><%=mobanshiyong%></p>
</div>
</li>
<%

  Next
%>
    </ul>
  </div>
</body>
</html>
<%	End If
	Set FsoFile = nothing
	Set Fso = nothing
%>

