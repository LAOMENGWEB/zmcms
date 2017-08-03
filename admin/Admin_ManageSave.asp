<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Conn.asp"-->
<!--#include file="Admin.asp"-->
<!--#include file="inc/md5.asp"-->
<!--#include file="inc/sha1.asp"-->
<!--#include file="inc/function.asp"-->
<%
password=replace(trim(Request("password")),"'","")
UserName=replace(trim(Request("UserName")),"'","")


if strLength(password)>16 or strLength(password)<6 then
 response.write  "<script language=javascript>alert('请输入的密码位数不能小于6位或大于16位!');history.go(-1);</script>"
  response.End
end if

'查找数据库，检查此管理员是否已经存在
set rs=server.createobject("adodb.recordset")
sql="select * from admin where UserName='" & UserName & "'"
rs.open sql,conn,1,1

if rs.recordcount >= 1 then
  response.write  "<script language=javascript>alert('此管理员帐号已经存在，请选用其他帐号!');history.go(-1);</script>"
  response.End
  rs.close
  set rs=nothing
end if

password=MD5(sha1(password&"zmcms"))
set rs=server.createobject("adodb.recordset")
sql="select * from admin"
rs.open sql,conn,1,3

'添加一个管理员帐号到数据库
rs.addnew
rs("UserName")=UserName
rs("PassWord")=password
if Request.Form("Working")=1 then
        rs("Working")=Request.Form("Working")
	  else
        rs("Working")=0
	  end if
	  
	  	  rs("AdminPurview")=Request.Form("Purview01") & Request.Form("Purview02") & Request.Form("Purview03") &_
	  
	                     Request.Form("Purview11") & Request.Form("Purview12") & Request.Form("Purview13") &_
	                     Request.Form("Purview21") & Request.Form("Purview22") & Request.Form("Purview23") & Request.Form("Purview24") &_
	                     Request.Form("Purview31") & Request.Form("Purview32") & Request.Form("Purview33") & Request.Form("Purview34") & Request.Form("Purview35") & Request.Form("Purview36") &_
						 Request.Form("Purview41") & Request.Form("Purview42") & Request.Form("Purview43") & Request.Form("Purview44") & Request.Form("Purview45") &_
	                     Request.Form("Purview51") & Request.Form("Purview52") & Request.Form("Purview53") &_
						 Request.Form("Purview511") & Request.Form("Purview512") & Request.Form("Purview513") & Request.Form("Purview514") & Request.Form("Purview515") &_
	                     Request.Form("Purview61") & Request.Form("Purview62") & Request.Form("Purview63") &_
						 	 Request.Form("Purview611") & Request.Form("Purview612") & Request.Form("Purview613") & Request.Form("Purview614") & Request.Form("Purview615") &_
	                     Request.Form("Purview71") & Request.Form("Purview72") & Request.Form("Purview73") &_
	                     Request.Form("Purview81") &_
	                     Request.Form("Purview91") &_
	                     Request.Form("Purview101") & Request.Form("Purview102") &_
	                     Request.Form("Purview103") &_
	                     Request.Form("Purview104") & Request.Form("Purview10411") & Request.Form("Purview10412") & Request.Form("Purview10413") & Request.Form("Purview10414") & Request.Form("Purview10415") &_
	                     Request.Form("Purview105") &_
	                     Request.Form("Purview106") & Request.Form("Purview107") &_
	                     Request.Form("Purview108") & Request.Form("Purview109") & Request.Form("Purview1091") & Request.Form("Purview1081") &_
	                     Request.Form("Purview110") &_
						 Request.Form("Purview1000001") & Request.Form("Purview1000002") & Request.Form("Purview1000003") & Request.Form("Purview1000004") &_
						 Request.Form("Purview111") &_
	                    Request.Form("Purview112") &_
	                    Request.Form("Purview113") & Request.Form("Purview1131") & Request.Form("Purview1132") &_
	                    Request.Form("Purview114") &_
	                     Request.Form("Purview201") & Request.Form("Purview202") &_
						 Request.Form("Purview1000") & Request.Form("Purview1001") & Request.Form("Purview1002") &_
						 Request.Form("Purview9999") &_
						 Request.Form("Purview301")
	  rs("Explain")=trim(Request.Form("Explain"))
	  rs("glyname")=trim(Request.Form("glyname"))
	  rs("Adddate")=now()
rs.update
rs.close
set rs=nothing
call addlog("添加管理员")
 response.Write "<script language=javascript>alert(""管理员添加成功"");window.location.href=""admin_manage.asp""</script>"
    response.End
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>zmcms_管理页面</title>
</head>

<body>
</body>
</html>
