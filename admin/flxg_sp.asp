<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="htmlconfig.asp"-->
<!--#include file="admin_html_function.asp"-->
<%
ID=request.QueryString("id")
ID1=request.QueryString("id1")
ID2=request.QueryString("id2")
ClassName=request("ClassName")
SmallImage=trim(request.form("SmallImage"))
mulu=request("mulu")
m=request("m")
px=request("px")
if request("Submit")="修改" then
	Set rs = Server.CreateObject("ADODB.Recordset")
	if ID<>"" THEN
	sql="select * from spclass where id="&ID
	END IF
	if ID1<>"" THEN
	sql="select * from spclass where id="&ID1
	END IF
	if ID2<>"" THEN
	sql="select * from spclass where id="&ID2
	END IF
	rs.open sql,conn,1,3
rs("SmallImage")=SmallImage
	rs("ClassName")=ClassName
	rs("mulu")=mulu
	rs("m")=m
    rs("px")=px
    rs("lang")=lang
	rs.update
	rs.close
	Set rs=nothing
	if yy=2 then
call htmltvfl 
call htmltv
   call htmll("","","default"&lang&".html","default.asp","m=",mde,"language=",lang,"","","","","","","","")
    call htmll("/"&video&"/","",""&tvclass&".html",""&video&"/index.asp","m=",mtv,"language=",lang,"","","","","","","","")
end if
	call addlog("修改"&shipin&"分类")
	response.Write "<script language=javascript>alert(""类别修改成功"");window.location.href=""fl_sp.asp""</script>"
    response.End
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>zmcms_管理页面</title>
<script language="JavaScript" src="js/jquery-1.7.2.min.js"></script>
<script type="text/javascript" src="../javascript/colorbox/jquery.colorbox-min.js"></script>
<link rel="stylesheet" href="../zmbjq/themes/default/default.css" />
<link rel="stylesheet" href="../zmbjq/plugins/code/prettify.css" />
<script charset="utf-8" src="../zmbjq/kindeditor-all.js"></script>
<script charset="utf-8" src="../zmbjq/lang/zh_CN.js"></script>
<script charset="utf-8" src="../zmbjq/plugins/code/prettify.js"></script>
<link href="../javascript/colorbox/colorbox.css" rel="stylesheet" media="screen">
<script language="JavaScript" src="js/ChinesePY.js"></script>
<script language="JavaScript" src="js/getpy.js"></script>
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
Set rs = Server.CreateObject("ADODB.Recordset")
if ID<>"" THEN
	sql="select * from spclass where id="&ID&" and lang='"&lang&"'"
	END IF
	if ID1<>"" THEN
	sql="select * from spclass where id="&ID1&" and lang='"&lang&"'"
	END IF
	if ID2<>"" THEN
	sql="select * from spclass where id="&ID2&" and lang='"&lang&"'"
	END IF
rs.open sql,conn,1,1
If rs.EOF Then
    response.Write "<script language=javascript>window.location.href=""fl_sp.asp""</script>"
    response.End
End if
%>
<form id="zmcms" name="zmcms" method="post" class="checkform">


<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center"> 修改<%=shipin%>分类 <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>


<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="shipin.asp" target="main">返回菜单</a></div><div class="dmeun"><a href="fl_sp.asp" target="main"><%=shipin%>分类管理</a></div></td>
 

 </tr>
</table>

  <table width="99%"  border="0" cellpadding="6" cellspacing="0" align="center" class="dtable4">
   
     <tr  >
				    <td >上级分类: </td><td>
			<% if request.QueryString("id")<>"" then%>
            <%
		 
		  set rsc=server.CreateObject("adodb.recordset")
		  rsc.open "select * from spclass where id="& request.QueryString("id")&"",conn,1,1
          response.Write("顶级分类")
          rsc.close
		set rsc=nothing
		
		  end if%>
		  		<% if request.QueryString("id1")<>"" then%>
            <%
		
		  set rsc=server.CreateObject("adodb.recordset")
		  rsc.open "select * from spclass where ID="& request.QueryString("id1")&"",conn,1,1
          ParentClassID=rsc("ParentClassID")
          rsc.close
		set rsc=nothing
		 
		  set rsc=server.CreateObject("adodb.recordset")
		  rsc.open "select * from spclass where id="& ParentClassID&"",conn,1,1
          response.Write("" & rsc("classname") & "")
          rsc.close
		set rsc=nothing
		
		  end if%>
		  <% if request.QueryString("id2")<>"" then%>
            <%
		
		  set rsc=server.CreateObject("adodb.recordset")
		  rsc.open "select * from spclass where ID="& request.QueryString("id2")&"",conn,1,1
          ParentClassID1=rsc("ParentClassID")
          rsc.close
		set rsc=nothing
		 
		  set rsc=server.CreateObject("adodb.recordset")
		  rsc.open "select * from spclass where id="& ParentClassID1&"",conn,1,1
          response.Write("" & rsc("classname") & "")
          rsc.close
		set rsc=nothing
	
		  end if%>
    </td>
			      </tr>
   
    <tr>
      <td>类别名称:</td><td>
      <input name="classname" type="text" id="classname" value="<%=rs("classname")%>" datatype="*" nullmsg="类别名称不能为空"/>

      </td>
    </tr>

			  <tr>
			    <td >分类图片：</td><td>
				  <input name="SmallImage"  type="text" id="weepic"  value="<%=rs("SmallImage")%>" size="50"  /><br/><br/>
			      <a class='sfile' href="Upload_All.Asp?add=one" title="支持上传格式(*.swf;*.gif;*.bmp)"><img src="Images/dtup.jpg"  border="0"/></a>
<div id="mysmaimg"><%if rs("Smallimage")<>"" then%>
<img src="../<%=rs("SmallImage")%>" id="sc"/><br />
<a href="delslt.asp?sltdz=<%=rs("Smallimage")%>" target="_blank" onClick="deltp()" id="lj">删除</a>
<%end if%>
</div>			        </td>
	      </tr>
          	 <tr>
      <td>目录名称:</td><td>
        <input name="mulu" type="text" value="<%=rs("mulu")%>" id="mulu" datatype="*" nullmsg="目录名称不能为空"/>  
		
		<input type="button" value="转换名称首字母" class="btnSearch" id="mulu_6" onClick="GetPy('', 'classname', 'mulu', 6, '', '');" />
<input type="button" value="转换名称全拼音" class="btnSearch" id="mulu_3" onClick="GetPy('', 'classname', 'mulu', 3, '', '');" />     
      * 生成html存放的路径  (最好以栏目名称的拼音缩写或者全拼 来命名。)</td>
    </tr>
              <tr>
      <td>高亮值:</td><td>
      <input name="m" type="text" id="m" value="<%=rs("m")%>"/>  
      高亮值 必须 和导航菜单的相关栏目 高亮值相同。
      </td>
      
    </tr> 
	 <tr>
      <td>排序:</td><td>
      <input name="px" type="text" id="px" value="<%=rs("px")%>" size="5" />
   
      </td>
    </tr>
	 <tr>
      <td colspan="2">
     <label>
      <input type="submit" name="Submit" value="修改" class="bnt"/>
      </label>
      </td>
    </tr>
  </table>
</form>
<%
rs.close
set rs=nothing
%>
</body>
</html>
