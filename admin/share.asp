<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="inc/function.asp"-->
<%
if Instr(session("AdminPurview"),"|103,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</script>"
  response.end
end if
'========判断是否具有管理权限
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>zmcms_管理页面</title>
<link href="css/style.css" rel="stylesheet" type="text/css">
	<script language="JavaScript" src="js/jquery-1.7.2.min.js"></script>
<style type="text/css">
table{ font-size: 14px;  border-right:none}
table td{padding:10px; font-family:"微软雅黑";color:#ccc; font-weight:bold;}
table td a{color:#ffffff}
</style>
<!--#include file="hqyy.asp"-->
</head>

<body>
<%
img=request("img")
top=Request("top")
wz=Request("wz")
jl=Request("jl")
If request.Form("submit") = "确认修改" Then
Set rs = server.CreateObject("adodb.recordset")
    sql = "select * from share "
    rs.Open sql, conn, 1, 3
    rs("img") = img
    rs("top") =top
	rs("wz") =wz
    rs("jl") =jl
    rs.update
    rs.Close
    Set rs = Nothing
    addlog("设置分享")
    response.Write "<script language=javascript>alert(""恭喜您修改成功！"");window.location.href=""share.asp""</script>"
    response.End
End If
%>
<form  name="mx" method="post" >
<%
Set rs = server.CreateObject("adodb.recordset")
sql = "select * from share where id=1"
rs.Open sql, conn, 1, 1
If rs.EOF Then
    response.Write "<script language=javascript>alert(""不存在此数据！"");window.history.back();</script>"
    response.End
End If
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" style=" background: #F9f9f9;border:none; ">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">分享参数设置  </div></td>
</tr>
</table>

<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="set.asp" target="main">返回菜单</a></div></td>
 

 </tr>
</table>
                      
                         <table width="99%"  border="0" cellpadding="0" cellspacing="0"  align="center" class="dtable4">
                                   
                                      <tr >
                                        <td  >按钮风格选择：</td>
										<td>
<input type="radio" name="img" id="img" value="0" <%if rs("img")=0 then response.write "checked"%> /> 
<img src="image/l0.gif" align="absmiddle" />
                                         
<input type="radio" name="img" id="img" value="1" <%if rs("img")=1 then response.write "checked"%> />
<img src="image/l1.gif" align="absmiddle" />

<input type="radio" name="img" id="img" value="2" <%if rs("img")=2 then response.write "checked"%> />
<img src="image/l2.gif" align="absmiddle" />
<input type="radio" name="img" id="img" value="3" <%if rs("img")=3 then response.write "checked"%>/>
<img src="image/l3.gif" align="absmiddle" />
<input type="radio" name="img" id="img" value="4" <%if rs("img")=4 then response.write "checked"%> />
<img src="image/l4.gif" align="absmiddle" />
<input type="radio" name="img" id="img" value="5" <%if rs("img")=5 then response.write "checked"%>/>
<img src="image/l5.gif" align="absmiddle" />
<input type="radio" name="img" id="img" value="6" <%if rs("img")=6 then response.write "checked"%>/>
<img src="image/l6.gif" align="absmiddle" />
<input type="radio" name="img" id="img" value="7" <%if rs("img")=7 then response.write "checked"%>/>
<img src="image/l7.gif" align="absmiddle" />
<input type="radio" name="img" id="img" value="8" <%if rs("img")=8 then response.write "checked"%>/>
<img src="image/l8.gif" align="absmiddle" /></td>
                                      </tr>
                                      <tr >
                                        <td  >离顶部距离：</td>                    <td>                      <input name="top" type="text" id="top" onMouseOver="style.backgroundColor='#FFF'"  value="<%=rs("top")%>" size="4" /> px</td>
                                      </tr>
                                      <tr >
                                        <td  >浮窗位置：</td>                    <td> 
                                          
                                          <input type="radio" name="wz" value="left" <%if rs("wz")="left" then response.write "checked"%> />
                                        左侧
                                        <input type="radio" name="wz" value="right" <%if rs("wz")="right" then response.write "checked"%> />
                                        右侧                                         </td>
                                      </tr>
                                      <tr >
                                        <td  >浮窗大小：</td>                    <td> 
                                          <input type="radio" name="jl" value="1" <%if rs("jl")=1 then response.write "checked"%> />
                                        一栏
                                        <input type="radio" name="jl" value="2" <%if rs("jl")=2 then response.write "checked"%> />
                                        二栏                                          </td>
                                      </tr>
                                     <tr >
                                        <td   >
  <input  class="bnt" type="submit" value="确认修改" id="submit4" name="submit" /></td>
                           </tr> 
   </table>

</form>
</body>
</html>
<%
  conn.Close
    Set conn = Nothing
%>