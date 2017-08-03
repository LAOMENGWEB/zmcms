<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="zmwbk.asp"-->
<!--#include file="sc.asp"-->
<!-- #include file="Inc/function.asp" -->
<%
if Instr(session("AdminPurview"),"|112,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</script>"
  response.end
end if
'========判断是否具有管理权限
%>
<%
If request("action") = "del" Then


pfpic=request.QueryString("pfpic")
call deletefile(pfpic)
Set rs = server.CreateObject("adodb.recordset")
sql = "select * from ad where lang='"&lang&"'  "
rs.Open sql, conn, 1, 3
rs("pfpic")=null
rs.update
    rs.Close
    Set rs = Nothing
   
  
	call addlog("删除漂浮广告图片")


end if
%>

<%
If request.Form("submit") = "图片提交" Then
pfpic=Trim(request.form("pfpic"))
Set rs = server.CreateObject("adodb.recordset")
rs.open "Select * From ad where lang='"&lang&"'" ,conn,1,3

 rs("pfpic")=pfpic
 rs("lang")=lang
rs.update
rs.close
set rs=nothing
end if
%>



<%
If request.Form("submit") = "确定修改" Then
pfpic=Trim(request.form("pfpic"))
pfurl=Trim(request.form("pfurl"))
pfwidth=Trim(request.form("pfwidth"))
pfheight=Trim(request.form("pfheight"))
pfxs=Trim(request.form("pfxs"))
Set rs = server.CreateObject("adodb.recordset")
rs.open "Select * From ad where lang='"&lang&"'" ,conn,1,3
rs("pfwidth")=pfwidth
rs("pfheight")=pfheight
 rs("pfpic")=pfpic
 rs("pfurl")=pfurl
rs.update
rs.close
set rs=nothing
call addlog("设置漂浮广告")
response.Write "<script language=javascript>alert(""恭喜您修改成功"");window.location.href=""plad.asp""</script>"
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>zmcms_管理页面</title>
<script language="JavaScript" src="js/jquery-1.7.2.min.js"></script>
 <script type="text/javascript" src="../js/Validform_v5.3.2_min.js"></script>
 <link href="css/style.css" rel="stylesheet" type="text/css">
	 <style type="text/css">
.table1{ font-size: 14px; background: #F9f9f9;border:none;}
.table1 td{padding:10px; font-family:"微软雅黑";color:#ccc; font-weight:bold;}
.table1 td a{color:#ffffff}
</style>
<!--#include file="hqyy.asp"-->
  <script type="text/javascript">
$(function(){
	
$(".checkform").Validform();  //就这一行代码！;
	
})
</script>
<script type="text/javascript">

function deltp()
{
document.getElementById("url3").value="";
}

 </script>
</head>

<body>
 <form   method="post" name="myform" >
          
        <%
Set rs = server.CreateObject("adodb.recordset")
sql = "select * from ad where lang='"&lang&"'"
rs.Open sql, conn, 1, 1
If rs.EOF Then
    response.Write "<script language=javascript>window.location.href=""chajian.asp""</script>"
    response.End
End If
%>  

<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">广告管理  <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>


<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td >
 <div class="dmeun"><a href="chajian.asp" target="main" style="color:#ffffff">返回菜单</a></div>
 
 
 </td>
 

 </tr>
</table>

<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td >

 <div class="dmeun" style=" background:#eeeeee; color:#999999"><a href="dladd.asp" target="main" style=" color:#999999">对联广告</a></div>
 <div class="dmeun" ><a href="plad.asp" target="main" style="color:#ffffff">漂浮广告</a></div>
 
 
 </td>
 

 </tr>
</table>
		
        <table width="99%" border="0" cellpadding="0" cellspacing="0"  align="center" class="dtable4">

  <tr >
				    <td  >图片地址: </td> <td><input name="pfpic" type="text" id="url3" size="30" value="<%=rs("pfpic")%>" />
<input type="button" id="image3" value="上传图片" /> <input type="submit" name="Submit" value="图片提交" /> 提示：先上传，在点击提交</br>
					<% if not IsFileExist(server.mappath(".."&sitepath&""&rs("pfpic")&"")) then 
		response.Write("<br /><font color=red>请上传您的广告图片</font>")   
		%>
			 <br />
		<%
		else
		%>
		<img src="../<%=rs("pfpic")%>" width="<%=rs("pfwidth")%>" hegiht="<%=rs("pfheight")%>"><br>
		<a href="?action=del&pfpic=<%=rs("pfpic")%>"><div style="width:80px;border:1px #ddd solid; background:#3399FF; text-align:center; color:#fff">删除</div></a>  注：上传新漂浮广告图片之前，请先删除以前的漂浮图片
		<%
		end if
		
		%>
		
					
					
					
					
					
					
					</td>
			      </tr>
				  <tr >
				    <td >图片链接:</td>
					<td>
		              <input name="pfurl" type="text" id="pfurl" value="<%=rs("pfurl")%>" /></td>
			      </tr>
				  <tr > 
                    <td > 广告尺寸:</td>
					<td>
                      <input name="pfwidth"type="text" class="inputtext" id="pfwidth" value="<%=rs("pfwidth")%>" size="6" maxlength="30" />
                    宽×高
                    <input name="pfheight"type="text" class="inputtext" id="pfheight" value="<%=rs("pfheight")%>" size="6" maxlength="30" /></td>
                  </tr>                 
                  <tr > 
                    <td colspan="2">  
                        
                        <input class="bnt" name="submit" type="submit" value="确定修改">
                        </td>
                  </tr>
            </table>
            
          </form>
       
</body>
</html>
