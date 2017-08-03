<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="htmlconfig.asp"-->
<!--#include file="admin_html_function.asp"-->
<%
ID=checknum(request.QueryString("ID"),0)
ClassName=request("ClassName")
mulu=request("mulu")
m=request("m")
SmallImage=trim(request.form("SmallImage"))
px=request("px")
if request("Submit")="提交" then

	Set rs = Server.CreateObject("ADODB.Recordset")
	sql="select * from cpclass Where  mulu= '" & mulu & "'"
	rs.open sql,conn,2,3

	rs.addnew
	rs("ClassName")=ClassName
	rs("SmallImage")=SmallImage
	rs("mulu")=mulu
	rs("px")=px
	rs("m")=m
	rs("xs")=1
	rs("lang")=lang
rs("ParentClassID")=ID
	rs.update
	rs.close
	Set rs=nothing
	
	ChildCount=conn.execute("select count(*) from [cpclass] where ParentClassID="&id)(0)
conn.execute("update [cpclass] set c_child = "& checknum(ChildCount,0) &" where id="& id &"")
	
	if yy=2 then
call htmlprofl
call htmlpro
  call htmll("","","default"&lang&".html","default.asp","m=",mde,"language=",lang,"","","","","","","","")
    call htmll("/"&product&"/","",""&Productclass&".html",""&product&"/index.asp","m=",mpro,"language=",lang,"","","","","","","","")
end if
	response.Write "<script language=javascript>alert("""&chanpin&"分类添加成功"");window.location.href=""fl_cp.asp""</script>"
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
<form id="zmcms" name="zmcms" method="post" class="checkform">



<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">增加<%=chanpin%>分类  <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>




<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="chanpin.asp" target="main">返回菜单</a></div><div class="dmeun"><a href="fl_cp.asp" target="main"><%=chanpin%>分类管理</a></div></td>
 

 </tr>
</table>


  <table width="99%"  border="0" cellpadding="0" cellspacing="0" align="center" class="dtable4">
    <tr>
				    <td  >上级分类: </td>
					<td>
			<% if request.QueryString("id")<>"" then%>
            <%
		 
		  set rsc=server.CreateObject("adodb.recordset")
		  rsc.open "select * from cpclass where id="& request.QueryString("id")&"",conn,1,1
          response.Write("" & rsc("classname") & "")
          rsc.close
		set rsc=nothing
		else
	
		  response.write("顶级分类")
		  end if%>
		  		
    </td>
			      </tr>
    <tr>
      <td>类别名称:</td>
					<td>
      <input name="classname" type="text" id="classname" datatype="*" nullmsg="类别名称不能为空"/>
  
      </td>
    </tr>
	 <tr >

			    <td >分类图片：</td>
					<td>
				 <input name="SmallImage"  type="text" id="weepic"  value="" size="50" /><br/><br/>
			      <a class='sfile' href="Upload_All.Asp?add=one" title="支持上传格式(*.swf;*.gif;*.bmp)"><img src="Images/dtup.jpg"  border="0"/></a>
		        <div id="mysmaimg"></div>		        </td>
	      </tr>
           <tr>
      <td>目录名称:</td>
					<td>
        <input name="mulu" type="text" id="mulu" datatype="*" nullmsg="目录名称不能为空"/> 
<input type="button" value="转换名称首字母" class="btnSearch" id="mulu_6" onClick="GetPy('', 'classname', 'mulu', 6, '', '');" />
<input type="button" value="转换名称全拼音" class="btnSearch" id="mulu_3" onClick="GetPy('', 'classname', 'mulu', 3, '', '');" /> 		     
      生成html存放的路径  (最好以栏目名称的拼音缩写或者全拼 来命名。)</td>
    </tr>
          <tr>
      <td>高亮值:</td><td>
      <input name="m" type="text" id="m"/>  
      高亮值 必须 和导航菜单的相关栏目 高亮值相同。
      </td>
     
    </tr> 
	 <tr>
      <td>排序:</td>
					<td>
      <input name="px" type="text" id="px" size="5" value="1"/>
   
      </td>
    </tr>
	 <tr>
      <td colspan="2">
     <label>
      <input type="submit" name="Submit" value="提交" class="bnt"/>
      </label>
      </td>
    </tr>
  </table>
</form>
</body>
</html>
