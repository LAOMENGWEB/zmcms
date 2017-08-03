<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="Inc/Function.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>zmcms_管理页面</title>
</head>

<body>
<%
if Instr(session("AdminPurview"),"|514,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</script>"
  response.end
end if
dim id
id=request.QueryString("id")
set rs=server.CreateObject("adodb.recordset")
sql = "select * from shoppingshipment where id = "&id&""
rs.open sql,conn,1,3
if rs.bof and rs.eof then
response.Write "<script language=javascript>alert(""不存在此数据！"");window.history.back();</script>"
response.End()
else
dim picUrl
dim picUrlArray
dim x,y
picUrl = getPicUrl(rs("content"))
if picUrl <> "" then
picUrl = left(picUrl,len(picUrl)-1)
picUrlArray = split(picUrl,",")
for x = 0 to ubound(picUrlArray)
if instr(picUrlArray(x),"") > 0 then
call DeleteFile(picUrlArray(x))
end if
next
end if
call deletefile1(SmallImage)
rs.delete
call addlog("删除配送方式")
response.Write "<script language=javascript>alert(""删除成功！"");window.location.href=""psfs.asp""</script>"
end if
rs.close

%>
</body>
</html>
