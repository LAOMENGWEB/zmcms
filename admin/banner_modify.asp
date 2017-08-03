<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!-- #include file="Inc/function.asp" -->
<%
If request.QueryString("id") = "" Or Not IsNumeric(request.QueryString("id")) Then
    response.Write "<script language=javascript>alert(""非法操作！"");window.history.back();</script>"
    response.End
End If




id=request.QueryString("id")



If request.Form("submit") = "确认修改" Then

    Set rs = server.CreateObject("adodb.recordset")
    sql = "select * from banner where id = "&request.QueryString("id")&""
    rs.Open sql, conn, 1, 3
    rs("banner_name") = request.Form("banner_name")
    rs("uploadfile") = request.Form("uploadfile")
    rs("banner_url") = request.Form("banner_url")
    rs("banner_order") = request.Form("banner_order")
	  rs("content") = request.Form("content")
	rs("zmone") = request.Form("zmone")
	rs("lang")=lang
    rs.update
    rs.Close
    Set rs = Nothing
   call addlog("修改幻灯片")
    response.Write "<script language=javascript>alert(""幻灯修改成功！"");window.location.href=""banner.asp""</script>"
    response.End
End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>zmcms_管理页面</title>

<script language="JavaScript" src="js/jquery-1.7.2.min.js"></script>
<script type="text/javascript" src="../javascript/colorbox/jquery.colorbox-min.js"></script>
 <script type="text/javascript" src="../js/Validform_v5.3.2_min.js"></script>
<link href="../javascript/colorbox/colorbox.css" rel="stylesheet" media="screen">
<link href="css/style.css" rel="stylesheet" type="text/css">
  <script type="text/javascript">
$(function(){
	
$(".checkform").Validform();  //就这一行代码！;
	
})
</script>
<script language=javascript>
$(document).ready(function(){
$(".allfile").colorbox({iframe: true, innerWidth: 1000,innerHeight: 500});
$(".sfile").colorbox({iframe: true, innerWidth: 400,innerHeight: 160});
});
</script>
<script type="text/javascript">

function deltp()
{
var parent=document.getElementById("mysmaimg");
var child=document.getElementById("sc");
var child1=document.getElementById("lj");
parent.removeChild(child);
parent.removeChild(child1);
document.getElementById("weepic").value="";
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


<%
Set rs = server.CreateObject("adodb.recordset")
sql = "select * from banner where id = "&request.QueryString("id")&" and lang='"&lang&"'"
rs.Open sql, conn, 1, 1
If rs.EOF Then
response.Write "<script language=javascript>window.location.href=""banner.asp""</script>"
response.End
End If
%>
        <form  name="myform" method="post" class="checkform">
<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">幻灯修改  <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>
<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="huandeng.asp" target="main">返回菜单</a></div>

<div class="dmeun"><a href="banner.asp" target="main">幻灯管理</a></div>
 </td>
 

 </tr>
</table>
<table width="99%"  border="0" cellpadding="0" cellspacing="0" align="center" class="dtable4">
		  <tr  >
                        <td >所属分类:</td>
						
						<td>
						<select name="zmone" id="select" datatype="*" nullmsg="所属分类不能为空">
      <%
		  dim rsc
		  set rsc=server.CreateObject("adodb.recordset")
		  rsc.open "select * from bannerclass",conn,1,1
		  Response.Write("<option value="""">请选择分类</option>") 
    If Rsc.Eof and Rsc.Bof Then
      Response.Write("<option value="""">请先添加分类</option>")
    Else
		  while not rsc.eof
		    if rs("zmone")=rsc("id") then
			response.Write("<option value=""" & rsc("id") & """ selected>" & rsc("classname") & "</option>")
			else
			response.Write("<option value=""" & rsc("id") & """>" & rsc("classname") & "</option>")
			end if
			rsc.movenext
		wend
		end if
		rsc.close
		set rsc=nothing
		  %>
                        </select></td>
                      </tr>
		 
		 
          <tr>
            <td  >文字说明：</td>
			<td>
              <input name="banner_name" type="text" value="<%=rs("banner_name")%>" size="80" /></td>
          </tr>
          <tr>
            <td  >链接地址：</td>
			<td>
            <input name="banner_url" type="text" value="<%=rs("banner_url")%>" size="80"/></td>
          </tr>
          

		   <tr >

			    <td >幻灯 图片：</td>
			<td>
				   <input name="uploadfile"  type="text" id="weepic"  value="<%=rs("uploadfile")%>" size="50"  /><br/><br/>
			      <a class='sfile' href="Upload_All.Asp?add=one" title="支持上传格式(*.swf;*.gif;*.bmp)"><div style="width:100px;padding:10px; background:#28b576;color:#ffffff" align="center">幻灯图片上传</div></a>
<div id="mysmaimg"><%if rs("uploadfile")<>"" then%>
<img src="../<%=rs("uploadfile")%>" id="sc"/><br />
<a href="delslt.asp?sltdz=<%=rs("uploadfile")%>" target="_blank" onClick="deltp()" id="lj">删除</a>
<%end if%>
</div>			        </td>

	      </tr>
		  <tr
>
            <td >内容提示：</td>
			<td>
        
              <textarea name="content" cols="100" rows="8"><%=rs("content")%></textarea>
        </td>
          </tr>
          
          <tr>
            <td  >排　　序：</td>
			<td>
            <input name="banner_order" type="text" value="<%=rs("banner_order")%>" size="10"/> 注：数字越小越靠前</td>
          </tr>
          <tr>
            <td  colspan="2">
              <input class="bnt" name="submit" type="submit" value="确认修改" />
     </td>
          </tr>
        </form>
</table>
</body>
</html>
<% conn.Close
    Set conn = Nothing%>