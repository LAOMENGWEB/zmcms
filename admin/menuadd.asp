<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="inc/function.asp"-->
<%
t=request.QueryString("t")
ID=checknum(request.QueryString("ID"),0)
ClassName=request("ClassName")
px=request("px")
mopen=request("mopen")
bieming=request("bieming")
url=request("url")
mcolor=request("mcolor")
bcolor=request("bcolor")
SmallImage=trim(request.form("SmallImage"))
htmlurl=request("htmlurl")
m=request("m")
if request("Submit")="提交" then
if ClassName="" then
response.Write "<script language=javascript>alert(""菜单名称不能为空"");window.history.back();</script>"
response.End
End If
If px = "" Or Not IsNumeric(px) Then
response.Write "<script language=javascript>alert(""排序名称不能为空！并且为数字"");window.history.back();</script>"
response.End
End If
Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from meun"
rs.open sql,conn,2,3
rs.addnew
rs("ClassName")=ClassName
rs("px")=px
rs("url")=url
rs("m")=m
rs("bieming")=bieming
rs("htmlurl")=htmlurl
rs("SmallImage")=SmallImage
rs("mopen")=mopen
rs("mcolor")=mcolor
rs("bcolor")=bcolor
rs("lang")=lang
if t=0 then
rs("fl")=0
end if
if t=1 then
rs("fl")=1
end if
if t=2 then
rs("fl")=2
end if
if t=3 then
rs("fl")=3
end if
rs("xs")=1
rs("ParentClassID")=ID
rs.update
rs.close
Set rs=nothing
ChildCount=conn.execute("select count(*) from [meun] where ParentClassID="&id)(0)
conn.execute("update [meun] set c_child = "& checknum(ChildCount,0) &" where id="& id &"")
if t=1 then
call addlog("添加顶部导航菜单")
response.Write "<script language=javascript>alert(""顶部菜单添加成功"");window.location.href=""menugl.asp?t=1""</script>"
end if
if t=0 then
call addlog("添加底部导航菜单")
response.Write "<script language=javascript>alert(""底部菜单添加成功"");window.location.href=""menugl.asp?t=0""</script>"
end if
if t=2 then
call addlog("添加其他导航菜单")
response.Write "<script language=javascript>alert(""其他菜单添加成功"");window.location.href=""menugl.asp?t=2""</script>"
end if
if t=3 then
call addlog("添加WAP导航菜单")
response.Write "<script language=javascript>alert(""WAP菜单添加成功"");window.location.href=""menugl.asp?t=3""</script>"
end if
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
<!--#include file="ys.asp"-->
 <script type="text/javascript" src="../js/Validform_v5.3.2_min.js"></script>
<link href="css/style.css" rel="stylesheet" type="text/css">
<link href="../javascript/colorbox/colorbox.css" rel="stylesheet" media="screen">
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
<form id="zmcms" name="zmcms" method="post" action="" class="checkform">


<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">
	 <%if t=1 then%>
添加顶部导航
 <%end if%>
 <%if t=0 then%>
添加底部导航
 <%end if%>
 <%if t=2 then%>
添加其他导航
 <%end if%>
 <%if t=3 then%>
添加手机导航
 <%end if%>  <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>



<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td >
	 <div class="dmeun"><a href="daohang.asp" target="main">返回菜单</a></div>
 <%if t=1 then%>
<div class="dmeun"><a href="menugl.asp?t=1" target="main">顶部导航管理</a></div>
 <%end if%>
 <%if t=0 then%>
<div class="dmeun"><a href="menugl.asp?t=0" target="main">底部导航管理</a></div>
 <%end if%>
 <%if t=2 then%>
<div class="dmeun"><a href="menugl.asp?t=2" target="main">其他导航管理</a></div>
 <%end if%>
 <%if t=3 then%>
<div class="dmeun"><a href="menugl.asp?t=3" target="main">手机导航管理</a></div>
 <%end if%>

 </td>
 

 </tr>
</table>
  <table width="99%"  border="0" cellpadding="0" cellspacing="0" align="center" class="dtable4">
   
    <tr>
      <td width="11%" >菜单名称:</td>
	  <td width="89%">
        <input name="classname" type="text" id="classname" datatype="*" nullmsg="菜单名称不能为空"/>      
      *  （菜单名称是必填项目）       </td>                      
    </tr>
	  <tr>
	  <td>菜单颜色:</td>
	  <td>
	  <input name="mcolor" type="text" id="color"   value="" size="8" /> 
	 
 <input   type="button" id="colorpicker" value="打开取色器" />  
        </td>
		
		</tr>
    
	 <%if t=3 then%>
	 <tr>
	   <td>菜单背景：</td>
	     <td>
	     <input name="bcolor" type="text" id="bcolor" />
     </td>
    </tr>
	
		 <tr>
	   <td>菜单图标：</td>
	     <td>
	     <input name="htmlurl" type="text" id="htmlurl" size="40" />	      </td>
    </tr>
	<%end if%>
	 <tr>
	   <td>别名：</td>
	     <td>
	  
	     <input name="bieming" type="text" id="bieming" />
      </td>
    </tr>
	
	 <tr >

			    <td >导航图片：</td>
	     <td>
				   <input name="SmallImage"  type="hidden" id="weepic"  value="" size="50"  />
			      <a class='sfile' href="Upload_All.Asp?add=one" title="支持上传格式(*.swf;*.gif;*.bmp)"><img src="Images/dtup.jpg"  border="0"/></a>
		        <div id="mysmaimg"></div>		        </td>

	      </tr>
	
	
	
	 <tr>
	   <td>动态链接地址：</td>
	     <td>
	     <input name="url" type="text" id="url" size="40" />
	      外部地址请完整填写包含“http://”    添加链接的时候如果需要高亮 m参数必须填写 值可以随意设置</td>
    </tr>
	<%if t=3 then%>
		 	<%else%>
	 <tr>
	   <td>静态链接地址：</td>
	     <td>
	     <input name="htmlurl" type="text" id="htmlurl" size="40" />	      </td>
    </tr>
	<%end if%>
       <tr >
	   <td>高亮值设置:</td>
	  <td>
	   
	   <input name="m" type="text" id="m" value="0" size="40" />
       高亮值 必须和 以上的m参数的值相同
       </td>
    </tr>
	 <tr>
	   <td>打开方式：</td>
	     <td>
              <select name="mopen" id="mopen">
                <option value="_self">原窗口</option>
                <option value="_blank">新窗口</option>
         </select></td>
    </tr>
	 <tr>
      <td>排序：</td>
	     <td>
      <input name="px" type="text" id="px" value="1" size="5"/>
      * 注：数字越小越靠前      </td>
    </tr>
	 
	 <tr>
      <td colspan="2">
   
      <input type="submit" name="Submit" value="提交" class="bnt" />
        </td>
    </tr>
  </table>
</form>
</body>
</html>
