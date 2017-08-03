<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="inc/function.asp"-->
<%
t=request.QueryString("t")
ID=request("ID")
ClassName=request("ClassName")
px=request("px")
mopen=request("mopen")
mcolor=request("mcolor")
bcolor=request("bcolor")
url=request("url")
m=request("m")
bieming=request("bieming")
SmallImage=trim(request.form("SmallImage"))
htmlurl=request("htmlurl")
xs=request("xs")
if request("Submit")="修改" then
Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from meun where id="&ID
rs.open sql,conn,1,3
rs("url")=url
rs("htmlurl")=htmlurl
rs("mopen")=mopen
rs("SmallImage")=SmallImage
rs("mcolor")=mcolor
rs("bcolor")=bcolor
rs("bieming")=bieming
rs("m")=m
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
rs("xs")=xs
rs("ClassName")=ClassName
rs("px")=px
rs.update
rs.close
Set rs=nothing
if t=1 then
call addlog("修改顶部导航菜单")
response.Write "<script language=javascript>alert(""顶部菜单修改成功"");window.location.href=""menugl.asp?t=1""</script>"
end if
if t=0 then
call addlog("修改底部导航菜单")
response.Write "<script language=javascript>alert(""底部菜单修改成功"");window.location.href=""menugl.asp?t=0""</script>"
end if
if t=2 then
call addlog("修改其他导航菜单")
response.Write "<script language=javascript>alert(""其他菜单修改成功"");window.location.href=""menugl.asp?t=2""</script>"
end if
if t=3 then
call addlog("修改wap导航菜单")
response.Write "<script language=javascript>alert(""wap菜单修改成功"");window.location.href=""menugl.asp?t=3""</script>"
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
<link href="../javascript/colorbox/colorbox.css" rel="stylesheet" media="screen">
<!--#include file="ys.asp"-->
 <script type="text/javascript" src="../js/Validform_v5.3.2_min.js"></script>
<link href="css/style.css" rel="stylesheet" type="text/css">
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

</style>
<!--#include file="hqyy.asp"-->
</head>

<body>
	
<%

Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from meun where id = "&ID&" and lang='"&lang&"'"
rs.open sql,conn,1,1
If rs.EOF Then
response.Write "<script language=javascript>window.location.href=""daohang.asp""</script>"

End If


%>
<form id="zmcms" name="zmcms" method="post" action="" class="checkform">




<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center"><%if t=1 then%>
修改顶部导航
 <%end if%>
 <%if t=0 then%>
修改底部导航
 <%end if%>
 <%if t=2 then%>
修改其他导航
 <%end if%>
 <%if t=3 then%>
修改手机导航
 <%end if%>  <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>




<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="daohang.asp" target="main">返回菜单</a></div></td>
 

 </tr>
</table>




  <table width="99%"  border="0" cellpadding="0" cellspacing="0"  align="center" class="dtable4">
   
    <tr >
      <td width="10%">菜单名称: </td>
	  <td width="90%">
      <input name="classname" type="text" id="classname" value="<%=rs("classname")%>" datatype="*" nullmsg="菜单名称不能为空"/>  *  菜单名称不能为空</td>
        
    </tr>
	  <tr >
	  <td>菜单颜色: </td>
	  <td>
	   <input name="mcolor" type="text" id="color"   value="<%=rs("mcolor")%>" size="8" /> 
 <input   type="button" id="colorpicker" value="打开取色器" />    </td>
    </tr>
	 <%if t=3 then%>
 <tr >
	   <td>菜单背景颜色:</td>
	  <td>
	 
	     <input name="bcolor" type="text" id="bcolor" value="<%=rs("bcolor")%>" />
   </td>
    </tr>
	
		 <tr>
	   <td>菜单图标:</td>
	  <td>
	     <input name="htmlurl" type="text" id="htmlurl" size="40" value="<%=rs("htmlurl")%>"/>	      </td>
    </tr>
	<%end if%>
	 <tr >
	   <td>别名:</td>
	  <td>
	 
	     <input name="bieming" type="text" id="bieming" value="<%=rs("bieming")%>" />
     </td>
    </tr>
	 <tr >
 
			    <td >导航图片：</td>
	  <td>
				   <input name="SmallImage"  type="hidden" id="weepic"  value="<%=rs("SmallImage")%>" size="50"  />
			      <a class='sfile' href="Upload_All.Asp?add=one" title="支持上传格式(*.swf;*.gif;*.bmp)"><img src="Images/dtup.jpg"  border="0"/></a>
<div id="mysmaimg"><%if rs("Smallimage")<>"" then%>
<img src="../<%=rs("SmallImage")%>" id="sc"/><br />
<a href="delslt.asp?sltdz=<%=rs("Smallimage")%>" target="_blank" onClick="deltp()" id="lj">删除</a>
<%end if%>
</div>			        </td>

	      </tr>
	
	
	
	
	
	
	 <tr >
	   <td>动态链接地址:</td>
	  <td>
	   
	   <input name="url" type="text" id="url" value="<%=rs("url")%>" size="40" />
        外部地址请完整填写包含"http://"    添加链接的时候如果需要高亮 m参数必须填写 值可以随意设置</td>
    </tr>
	<%if t=3 then%>
	<%else%>
	 <tr >
	   <td>静态链接地址:</td>
	  <td>
	   
	   <input name="htmlurl" type="text" id="htmlurl" value="<%=rs("htmlurl")%>" size="40" /></td>
    </tr>
	<%end if%>
    
     <tr >
	   <td>高亮值设置:</td>
	  <td>
	   
	   <input name="m" type="text" id="m" value="<%=rs("m")%>" size="40" />
       高亮值 必须和 以上的m参数的值相同
       </td>
    </tr>
    
    
	 <tr >
	   <td>打开方式：</td>
	  <td>
             <select name="mopen">
                <option value="_self"<%If rs("mopen") = "_self" Then%>selected="selected"<%End If%>>原窗口</option>
                <option value="_blank"<%If rs("mopen") = "_blank" Then%>selected="selected"<%End If%>>新窗口</option>
            </select></td>
    </tr>
	 <tr >
	   <td>是否显示：</td>
	  <td>
	   <select name="xs">
                <option value="true"<%If rs("xs") = True Then%>selected="selected"<%End If%>>显示</option>
                <option value="false"<%If rs("xs") = false Then%>selected="selected"<%End If%>>隐藏</option>
            </select></td>
    </tr>
	 <tr >
      <td>排序:</td>
	  <td>
      <input name="px" type="text" id="px" value="<%=rs("px")%>" size="5" />
      *   注：数字越小越靠前   </td>
    </tr>
	 <tr >
      <td colspan="2">
  
      <input type="submit" class="bnt"name="Submit" value="修改" />
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
