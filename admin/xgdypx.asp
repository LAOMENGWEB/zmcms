<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="inc/Function.asp"-->
<!--#include file="htmlconfig.asp"-->
<!--#include file="admin_html_function.asp"-->
<style type="text/css">

body{ background:#f9f9f9}
</style>
<%
If request.Form("submit") = "修改" Then
px=trim(request("px"))
Set rs = Server.CreateObject("ADODB.Recordset")
sql1="select * from dy where id="&request.QueryString("id")&""
rs.open sql1,conn,1,3

  rs("px")=px
 rs.update
rs.close
set rs=nothing
if yy=2 then
call htmlabout
call htmll("","","default"&lang&".html","default.asp","m=",mde,"language=",lang,"","","","","","","","")
end if
call addlog("设置单页排序")
end if
Set rs = server.CreateObject("adodb.recordset")
sql2 = "select * from dy where id = "&request.QueryString("id")&""
rs.Open sql2, conn, 1, 1
If rs.EOF Then
    response.Write "<script language=javascript>alert(""不存在此数据！"");window.history.back();</script>"
    response.End
End If
%>
<form id="form1" name="form1" method="post" action="xgdypx.asp?id=<%=rs("id")%>">
  <label>
  <input name="px" type="text" id="px" value="<%=rs("px")%>" size="5" style="border:1px #ddd solid;width:35px;height:20px; line-height:20px;padding-left:2px"/>
  </label>

  <label>
  <input type="submit" name="Submit" value="修改"  style="background:none;border:none;width:40px;height:20px;line-height:20px;"/>
  </label>
</form>
