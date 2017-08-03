<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="inc/function.asp"-->
<%
if Instr(session("AdminPurview"),"|613,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</script>"
  response.end
end if
'========判断是否具有管理权限
%>
<% 
t=request.QueryString("t")
function HtmlStrReplace(Str)'写入Html网页替换字符
  if Str="" or isnull(Str) then 
    HtmlStrReplace=""
    exit function 
  else
    HtmlStrReplace=replace(Str,"&lt;br&gt;","<br>")'"<br>"
  end if
end function
Function RndNumber(Min,Max) 
Randomize 
RndNumber=Int((Max - Min + 1) * Rnd() + Min) 
End Function
If request.QueryString("id") = "" Or Not IsNumeric(request.QueryString("id")) Then
    response.Write "<script language=javascript>alert(""非法操作！"");window.history.back();</script>"
    response.End
End If
tagcontent=HtmlStrReplace(request("tagcontent"))
tagname=request("tagname")
fl=request("fl")

If request.Form("submit") = "确认修改" Then
Set rs = server.CreateObject("adodb.recordset")
sql = "select * from templateTags where id = "&request.QueryString("id")&" "
rs.Open sql, conn3, 1, 3
rs("tagname") = tagname
rs("tagcontent") =tagcontent
rs("fl") = fl
    rs.update
    rs.Close
    Set rs = Nothing
  call addlog("修改标签样式")
      if t=1 then
    response.Write "<script language=javascript>alert(""恭喜您修改成功！"");window.location.href=""tagstyle.asp?t=1""</script>"
	end if
	 if t=2 then
    response.Write "<script language=javascript>alert(""恭喜您修改成功！"");window.location.href=""tagstyle.asp?t=2""</script>"
	end if
    response.End
End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>zmcms_管理页面</title>
<script language="JavaScript" src="js/jquery-1.7.2.min.js"></script>
 <script type="text/javascript" src="../js/Validform_v5.3.2_min.js"></script>
 <link href="css/style.css" rel="stylesheet" type="text/css">
  <script type="text/javascript">
$(function(){
	
$(".checkform").Validform();  //就这一行代码！;
	
})
</script>
<style type="text/css">
.table1{ font-size: 14px; background: #F9f9f9;border:none;}
.table1 td{padding:10px; font-family:"微软雅黑";color:#ccc; font-weight:bold;}
.table1 td a{color:#ffffff}
</style>
<!--#include file="hqyy.asp"-->
</head>

<body>

<form  name="mx" method="post" class="checkform">

<%
Set rs = server.CreateObject("adodb.recordset")
sql = "select * from templateTags where id = "&request.QueryString("id")&""
rs.Open sql, conn3, 1, 1
If rs.EOF Then
response.Write "<script language=javascript>alert(""不存在此数据！"");window.history.back();</script>"
response.End
End If
%>




<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center"> <%if t=1 then%>修改PC端标签样式 <%end if%>
 <%if t=2 then%>修改WAP端标签样式 <%end if%></div></td>
</tr>
</table>

<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="biaoqian.asp" target="main">返回菜单</a></div></td>
 

 </tr>
</table>
                     <table width="99%" border="0" cellpadding="0" cellspacing="0"  align="center" class="dtable4">
                                     
                         <tr >
                                        <td   >所属分类：</td>
										<td>
										<select name="fl" id="fl">
										  <option value="1" <%if rs("fl")=1 then Response.Write("selected")%>>PC端</option>
										  <option value="2" <%if rs("fl")=2 then Response.Write("selected")%>>WAP端</option>
										</select>
										</td>
                                      </tr>
                                      <tr >
                                        <td   >标签名称：</td>
										<td>
                                         <input type="text" name="tagname" value="<%=rs("tagname")%>" class="text-long" datatype="*" nullmsg="标签名称不能为空"/></td>
                                      </tr>
                                      <tr >
                                        <td   >显示样式代码：</td>
										<td>
                                        <textarea rows="20" name="tagcontent" id="tagcontent" cols="100" datatype="*" nullmsg="代码不能为空"><%=rs("tagcontent")%></textarea></td>
                                      </tr>
                                    
                                      <tr >
                                        <td colspan="2">
                                                                                <div align="left">
                                                                                  <input class="bnt" type="submit" value="确认修改" id="submit4" name="submit" />
                                                                                </div></td></tr>
   </table>
                        
</form>
<% 
Set rs = Nothing
conn3.Close
Set conn3 = Nothing
%>
</body>
</html>
