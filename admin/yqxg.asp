<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!-- #include file="Inc/function.asp" -->
<%
If request.QueryString("id") = "" Or Not IsNumeric(request.QueryString("id")) Then
    response.Write "<script language=javascript>alert(""非法操作！"");window.history.back();</script>"
    response.End
End If

If request("action") = "del" Then

id=request.QueryString("id")
logourl=request.QueryString("logourl")
call deletefile(logourl)
Set rs = server.CreateObject("adodb.recordset")
sql = "select * from yqlj where id="&id&" "
rs.Open sql, conn, 1, 3
rs("logourl")=null
rs.update
    rs.Close
    Set rs = Nothing
   
  
	call addlog("删除链接图片")


end if
If request.Form("submit") = "确认修改" Then

 
    Set rs = server.CreateObject("adodb.recordset")
    sql = "select * from yqlj where id = "&request.QueryString("id")&""
    rs.Open sql, conn, 1, 3
   
    rs("zmone") = request.Form("zmone")
     rs("lx") = request.Form("lx")
    rs("mc") = request.Form("mc")
    rs("logourl") = request.Form("logourl")
    rs("dz") = request.Form("dz")
	 rs("px") = request.Form("px")
	 rs("lang")=lang
    rs.update
    rs.Close
    Set rs = Nothing
   call addlog("修改链接")
    response.Write "<script language=javascript>alert(""链接修改成功！"");window.location.href=""yq.asp""</script>"
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
<link href="../javascript/colorbox/colorbox.css" rel="stylesheet" media="screen">
<script type="text/javascript" src="../js/Validform_v5.3.2_min.js"></script>
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


<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">修改链接  <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>


<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="lianjie.asp" target="main">返回菜单</a></div><div class="dmeun"><a href="yq.asp" target="main">链接管理</a></div></td>
 

 </tr>
</table>
  <form  name="yqlj" method="post" class="checkform">
 <table width="99%" border="0"  cellpadding="0" cellspacing="0"  align="center" class="dtable4">
      				<%
Set rs = server.CreateObject("adodb.recordset")
sql = "select * from yqlj where id = "&request.QueryString("id")&" and lang='"&lang&"'"
rs.Open sql, conn, 1, 1
If rs.EOF Then
    response.Write "<script language=javascript>window.location.href=""yq.asp""</script>"
    response.End
End If
%>
      
		
                         <tr  >
                        <td width="11%">所属分类:</td><td width="89%"><select name="zmone" id="select" datatype="*" nullmsg="所属分类不能为空">
      <%
		  dim rsc
		  set rsc=server.CreateObject("adodb.recordset")
		  rsc.open "select * from yqljclass",conn,1,1
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
            <td >链接名称</td><td>
              <input  name="mc" type="text" id="mc" value="<%=rs("mc")%>" size="15" datatype="*" nullmsg="链接名称不能为空"/></td>
          </tr>
          <tr>
            <td >链接地址：</td><td>
            <input  name="dz" type="text" id="dz" value="<%=rs("dz")%>" size="80" datatype="*" nullmsg="链接地址不能为空"/></td>
          </tr>
		  <tr >

			    <td >链接 图片：</td><td>
				   <input name="logourl"  type="text" id="weepic"  value="<%=rs("logourl")%>" size="50"  /><br/><br/>
			      <a class='sfile' href="Upload_All.Asp?add=one" title="支持上传格式(*.swf;*.gif;*.bmp)"><div style="width:100px;padding:10px; background:#28b576;color:#ffffff" align="center">logo上传</div></a><br/><br/>
<div id="mysmaimg"><%if rs("logourl")<>"" then%>
<img src="../<%=rs("logourl")%>" id="sc"/><br />
<a href="delslt.asp?sltdz=<%=rs("logourl")%>" target="_blank" onClick="deltp()" id="lj">删除</a>
<%end if%>
</div>			        </td>

	      </tr>
  <tr>
            <td >排序：</td><td>
            <input  name="px" type="text" id="px" value="<%=rs("px")%>" size="4"/></td>
          </tr>
      
          <tr>
            <td colspan="2" >
              <input class="bnt"  name="submit" type="submit" value="确认修改" />
            </td>
          </tr>
        </form>
</table>
</body>
</html>
<% conn.Close
    Set conn = Nothing%>