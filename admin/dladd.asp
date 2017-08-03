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


pic=request.QueryString("pic")
call deletefile(pic)
Set rs = server.CreateObject("adodb.recordset")
sql = "select * from ad where lang='"&lang&"' "
rs.Open sql, conn, 1, 3
rs("pic")=null
rs.update
    rs.Close
    Set rs = Nothing
   
  
	call addlog("删除对联左侧广告图片")


end if
%>
<%
If request("action") = "del1" Then


pic=request.QueryString("pic1")
call deletefile(pic1)
Set rs = server.CreateObject("adodb.recordset")
sql = "select * from ad where lang='"&lang&"'  "
rs.Open sql, conn, 1, 3
rs("pic1")=null
rs.update
    rs.Close
    Set rs = Nothing
   
  
	call addlog("删除对联左侧广告图片")


end if
%>

<%
If request.Form("submit") = "左侧图片提交" Then
pic=Trim(request.form("pic"))
Set rs = server.CreateObject("adodb.recordset")
rs.open "Select * From ad where lang='"&lang&"'" ,conn,1,3

 rs("pic")=pic

rs.update
rs.close
set rs=nothing
end if
%>

<%
If request.Form("submit2") = "右侧图片提交" Then
pic1=Trim(request.form("pic1"))
Set rs = server.CreateObject("adodb.recordset")
rs.open "Select * From ad where lang='"&lang&"'" ,conn,1,3

 rs("pic1")=pic1

rs.update
rs.close
set rs=nothing
end if
%>






<%
If request.Form("submit") = "确定修改" Then
pic=Trim(request.form("pic"))
url=Trim(request.form("url"))
width=Trim(request.form("width"))
height=Trim(request.form("height"))
zx=Trim(request.form("zx"))
top=Trim(request.form("top"))
left1=Trim(request.form("left1"))
right1=Trim(request.form("right1"))
pic1=Trim(request.form("pic1"))
url1=Trim(request.form("url1"))
lf=Trim(request.form("lf"))
rf=Trim(request.form("rf"))
Set rs = server.CreateObject("adodb.recordset")
rs.open "Select * From ad where lang='"&lang&"'" ,conn,1,3
rs("width")=width
rs("height")=height
 rs("pic")=pic
 rs("url")=url
 rs("top")=top
 rs("left1")=left1
 rs("right1")=right1
 rs("pic1")=pic1
 rs("url1")=url1
 rs("lf")=lf
 rs("rf")=rf
rs.update
rs.close
set rs=nothing
call addlog("设置对联广告")
response.Write "<script language=javascript>alert(""恭喜您修改成功"");window.location.href=""dladd.asp""</script>"
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
 
 <script type="text/javascript">

function deltpp()
{
document.getElementById("url2").value="";

}

 </script>

<style type="text/css">
.table1{ font-size: 14px; background: #F9f9f9;border:none;}
.table1 td{padding:10px; font-family:"微软雅黑";color:#ccc; font-weight:bold;}
.table1 td a{color:#ffffff}
</style>
<!--#include file="hqyy.asp"-->
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

 <div class="dmeun" ><a href="dladd.asp" target="main" style="color:#ffffff">对联广告</a></div>
 <div class="dmeun" style=" background:#eeeeee; color:#999999"><a href="plad.asp" target="main" style=" color:#999999">漂浮广告</a></div>
 
 
 </td>
 

 </tr>
</table>
		
		<table width="99%" border="0" cellpadding="0" cellspacing="0"align="center" class="dtable4">

				
				  
  <tr>
				    <td width="18%" >对联广告左侧图片地址: </td>
					<td width="82%"><input name="pic" type="text" id="url3" size="30" value="<%=rs("pic")%>" />
<input type="button" id="image3" value="上传图片" /> <input type="submit" name="Submit" value="左侧图片提交" /> 提示：先上传，在点击提交</br>				<% if not IsFileExist(server.mappath(".."&sitepath&""&rs("pic")&"")) then 

		response.Write("<br /><font color=red>请上传您的左侧广告图片</font>")   
		%>
			
			 <br />
		<%
		else
		%>
		<img src="../<%=rs("pic")%>" width="<%=rs("width")%>" hegiht="<%=rs("height")%>"><br>
		<a href="?action=del&pic=<%=rs("pic")%>" onClick="deltp()"><div style="width:80px;border:1px #ddd solid; background:#3399FF; text-align:center; color:#fff">删除</div></a>  注：上传新图片之前，请先删除以前的图片
	  <%
		end if
		
		%></td>
	      </tr>
				  <tr>
				    <td  >对联广告左侧网站地址:</td><td>
			          <input name="url" type="text" id="url" value="<%=rs("url")%>" /></td>
			      </tr>
				  <tr>
				    <td >对联广告右侧图片地址：</td><td><input name="pic1" type="text" id="url2" size="30" value="<%=rs("pic1")%>" />
<input type="button" id="image2" value="上传图片" />  <input type="submit" name="Submit2" value="右侧图片提交" /><br /></br><% if not IsFileExist(server.mappath(".."&sitepath&""&rs("pic1")&"")) then 
		response.Write("<font color=red>请上传您的右侧广告图片</font>")   
		%>
			
		<%
		else
		%>
		<img src="../<%=rs("pic1")%>" width="<%=rs("width")%>" hegiht="<%=rs("height")%>"><br>
		<a href="?action=del1&pic1=<%=rs("pic1")%>"  onclick="deltpp()"><div style="width:80px;border:1px #ddd solid; background:#3399FF; text-align:center; color:#fff">删除</div></a>  注：上传新图片之前，请先删除以前的图片
		<%
		end if
		
		%></br>				      </td>
			      </tr>
				  <tr>
				    <td >对联广告右侧图片链接：</td><td>
				      <label>
				      <input name="url1" type="text" id="url1" value="<%=rs("url1")%>" />
				      </label></td>
			      </tr>
				  <tr> 
                    <td > 广告尺寸:</td><td>
                      <input name="width"type="text" class="inputtext" id="width" value="<%=rs("width")%>" size="6" maxlength="30" />
                    宽×高
                    <input name="height"type="text" class="inputtext" id="height" value="<%=rs("height")%>" size="6" maxlength="30" /></td>
                  </tr>
				  <tr>
				    <td  >对联广告离顶部，左侧，右侧的距离</br></br></td><td>顶部:
                      <input name="top"type="text" class="inputtext" id="top" value="<%=rs("top")%>" size="6" maxlength="30" />
				      左侧:
				      <input name="left1"type="text" class="inputtext" id="left1" value="<%=rs("left1")%>" size="6" maxlength="30" />
				      右侧:
			        <input name="right1"type="text" class="inputtext" id="right1" value="<%=rs("right1")%>" size="6" maxlength="30" /></td>
			      </tr>
			
                  
                  
                 
                  <tr>
                    <td>左侧对联是否显示：</td>
					<td>
                      <label>
                    <select name="lf" id="lf">
                        <option value="block" <%if rs("lf")="block" then Response.Write("selected")%>>是</option>
                        <option value="none" <%if rs("lf")="none" then Response.Write("selected")%>>否</option>
                    </select>
                    </label></td>
                  </tr>
                  <tr>
                    <td>右侧对联是否显示:</td>   <td><label>
                      <select name="rf" id="rf">
                        <option value="block" <%if rs("rf")="block" then Response.Write("selected")%>>是</option>
                        <option value="none" <%if rs("rf")="none" then Response.Write("selected")%>>否</option>
                    </select>
                    </label></td>
                  </tr>
                  <tr> 
                    <td colspan="2"> 
                        
                        <input class="bnt"name="submit" type="submit" value="确定修改">                      </td>
                  </tr>
            </table>
            
</form>
</body>
</html>
