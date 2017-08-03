<!--#include file="../base.asp"-->
<!--#include file="../../lib/db.asp"-->
<!--#include file="../../lib/function.asp"-->
<!--#include file="../../lib/weixin_class.asp"-->
<!--#include file="../../lib/md5.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>删除自定义菜单</title>
</head>

<body>

<%
dim id
id=request.QueryString("id")
set rs=server.CreateObject("adodb.recordset")
sql = "select * from Plug_PicMsg where p_id = "&id&""
rs.open sql,conn2,1,3
p_Pic=rs("p_Pic")

if rs.bof and rs.eof then
response.Write "<script language=javascript>alert(""不存在此数据！"");window.history.back();</script>"
response.End()
end if
dim picUrl
dim picUrlArray
dim x,y
picUrl = getPicUrl(rs("P_content"))
if picUrl <> "" then
picUrl = left(picUrl,len(picUrl)-1)
picUrlArray = split(picUrl,",")
for x = 0 to ubound(picUrlArray)
if instr(picUrlArray(x),"") > 0 then
call DeleteFile(picUrlArray(x))
end if
next
end if

dim picUrl1
dim picUrlArray1
dim x1,y1
picUrl1 = rs("ImagePath")
if picUrl1 <> "" then
picUrlArray1 = split(picUrl1,"|")
for x1 = 0 to ubound(picUrlArray1)
DeleteFile(picUrlArray1(x1))
next
end if

call deletefile1(p_Pic)
conn2.Execute ( "delete from sys_KeyWord where k_plugName='picMsg' and k_plugParam = "&id&"" )

conn2.Execute ( "delete from Plug_PicMsg where P_id ="&id&"" )
rs.close
set rs=nothing

response.Write "<script language=javascript>alert(""删除成功！"");window.location.href=""admin.asp""</script>"

%>

</body>
</html>
