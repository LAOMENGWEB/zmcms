<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!-- #include file="admin.asp" -->
<!-- #include file="Inc/function.asp" -->


<%
if Instr(session("AdminPurview"),"|615,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</script>"
  response.end
end if
'========判断是否具有管理权限
%>




<%  
t=request.QueryString("t")

If request.Form("submit") = "保存" Then
if t=1 then
waplogo=request("waplogo")
waplogowidth=request("waplogowidth")
waplogoheight=request("waplogoheight")
end if

if t=2 then
wapnewinfo=request("wapnewinfo")
wapProductInfo=request("wapProductInfo")
wapalinfo=request("wapalinfo")
wapdowninfo=request("wapdowninfo")
waptvinfo=request("waptvinfo")
wapzpinfo=request("wapzpinfo")
waplyinfo=request("waplyinfo")
end if
if t=3 then
wapmap=request("wapmap")
wapmc=request("wapmc")
mapphone=request("mapphone")
wapcontent=request("wapcontent")
end if
if t=4 then
ewm=request("ewm")
ewmwidth=request("ewmwidth")
ewmheight=request("ewmheight")
end if
if t=5 then
wxewm=request("wxewm")
wxewmwidth=request("wxewmwidth")
wxewmheight=request("wxewmheight")
end if
Set rs = server.CreateObject("adodb.recordset")
sql = "select * from wzset where lang='"&lang&"' "
rs.Open sql, conn, 1, 3
if t=1 then
rs("waplogo")=waplogo
rs("waplogowidth")=waplogowidth
rs("waplogoheight")=waplogoheight
end if
if t=2 then
rs("wapnewinfo")=wapnewinfo
rs("wapProductInfo")=wapProductInfo
rs("wapalinfo")=wapalinfo
rs("wapdowninfo")=wapdowninfo
rs("waptvinfo")=waptvinfo
rs("wapzpinfo")=wapzpinfo
rs("waplyinfo")=waplyinfo
end if
if t=3 then
rs("wapmap")=wapmap
rs("wapmc")=wapmc
rs("mapphone")=mapphone
rs("wapcontent")=replace(request("wapcontent"),vbCRLF,"")
end if
if t=4 then
rs("ewm")=ewm
rs("ewmwidth")=ewmwidth
rs("ewmheight")=ewmheight
end if
if t=5 then
rs("wxewm")=wxewm
rs("wxewmwidth")=wxewmwidth
rs("wxewmheight")=wxewmheight
end if
rs.update
    rs.Close
    Set rs = Nothing
   if t=1 then
	call addlog("修改wap基本设置")
  response.Write "<script language=javascript>alert(""恭喜您修改成功"");window.location.href=""wap.asp?t=1""</script>"
    response.End
	end if
	  if t=2 then
	call addlog("修改wap列表页数量")
  response.Write "<script language=javascript>alert(""恭喜您修改成功"");window.location.href=""wap.asp?t=2""</script>"
    response.End
	end if
	  if t=3 then
	call addlog("修改wap地图信息")
  response.Write "<script language=javascript>alert(""恭喜您修改成功"");window.location.href=""wap.asp?t=3""</script>"
    response.End
	end if
	  if t=4 then
	call addlog("修改wap二维码")
  response.Write "<script language=javascript>alert(""恭喜您修改成功"");window.location.href=""wap.asp?t=4""</script>"
    response.End
	end if
	  if t=5 then
	call addlog("修改微信二维码")
  response.Write "<script language=javascript>alert(""恭喜您修改成功"");window.location.href=""wap.asp?t=5""</script>"
    response.End
	end if
End If
%>


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>zmcms_管理页面</title>

<script language="javascript" src="../js/Admin.js"></script>
<script type="text/javascript" src="js/jquery-1.7.2.min.js"></script>
<script type="text/javascript" src="../javascript/colorbox/jquery.colorbox-min.js"></script>
<link href="../javascript/colorbox/colorbox.css" rel="stylesheet" media="screen">
<link rel="stylesheet" href="../zmbjq/themes/default/default.css" />
<link rel="stylesheet" href="../zmbjq/plugins/code/prettify.css" />
<script charset="utf-8" src="../zmbjq/kindeditor-all.js"></script>
<script charset="utf-8" src="../zmbjq/lang/zh_CN.js"></script>
<script charset="utf-8" src="../zmbjq/plugins/code/prettify.js"></script>
<script language=javascript>
$(document).ready(function(){
$(".allfile").colorbox({iframe: true, innerWidth: 1000,innerHeight: 500});
$(".sfile").colorbox({iframe: true, innerWidth: 400,innerHeight: 160});
});
</script>

<link href="css/style.css" rel="stylesheet" type="text/css">
<style type="text/css">
.table1{ font-size: 14px; background: #F9f9f9;border:none;}
.table1 td{padding:10px; font-family:"微软雅黑";color:#ccc; font-weight:bold;}
.table1 td a{color:#ffffff}
</style>
<!--#include file="hqyy.asp"-->
<%if t=1 or t=4 or t=5 then%>
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
<%end if%>
</head>


<body>
<%if t=1 then%>
<form name="editForm" method="POST" action="wap.asp?t=1"  >
<%
end if
%>
<%if t=2 then%>
<form name="editForm" method="POST" action="wap.asp?t=2"  >
<%
end if
%>
<%if t=3 then%>
<form name="editForm" method="POST" action="wap.asp?t=3"  >
<%
end if
%>
<%if t=4 then%>
<form name="editForm" method="POST" action="wap.asp?t=4"  >
<%
end if
%>
<%if t=5 then%>
<form name="editForm" method="POST" action="wap.asp?t=5"  >
<%
end if
%>
<%
Set rs = server.CreateObject("adodb.recordset")
sql = "select * from wzset where lang='"&lang&"'"
rs.Open sql, conn, 1, 1



%>



<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center"><%if t=1 then%>手机版基本设置<%end if%>
 <%if t=2 then%>手机版列表页数量<%end if%>
 <%if t=3 then%>手机版地图设置<%end if%>
 <%if t=4 then%>手机版手机二维码上传<%end if%>
 <%if t=5 then%>手机版微信二维码上传<%end if%>  <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>



<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="shouji.asp" target="main">返回菜单</a></div></td>
 

 </tr>
</table>
<table width="99%" border="0" align="center" cellpadding="0" cellspacing="0"  class="dtable4">
  

   
 <%if t=1 then%>
      <tr ><td width="18%" >手机LOGO上传：</td>
  <td width="82%" >
				   <input name="waplogo"  type="text" id="weepic"  value="<%=rs("waplogo")%>" size="50"  /><br/><br/>
			      <a class='sfile' href="Upload_All.Asp?add=one" title="支持上传格式(*.swf;*.gif;*.bmp)"><div style="width:100px;padding:10px; background:#28b576;color:#ffffff" align="center">点此上传</div></a>
		        <div id="mysmaimg"><%if rs("waplogo")<>"" then%>
<img src="../<%=rs("waplogo")%>" id="sc"/><br />
<a href="delslt.asp?sltdz=<%=rs("waplogo")%>" target="_blank" onClick="deltp()" id="lj">删除</a>
<%end if%></div>		        </td>
</tr>
    </tr>
		
<tr> <td>LOGO大小：</td>
  <td>
	  
         <input name="waplogowidth"type="text" class="inputtext" id="waplogowidth"  value="<%=rs("waplogowidth")%>" size="6" maxlength="30" />
                    waplogo宽×高
        <input name="waplogoheight"type="text" class="inputtext" id="waplogoheight"  value="<%=rs("waplogoheight")%>" size="6" maxlength="30" /></td>
    </tr>
	<%end if%>
       <%if t=2 then%>
      <tr>
          <td width="18%">WAP新闻列表页调用条数：</td>
		   <td > <input name="wapnewinfo" type="text" id="wapnewinfo" value="<%=rs("wapnewinfo")%>" size="3"></td>
    </tr>
        <tr>
          <td >WAP产品列表页调用条数：</td>
		   <td >
		    <input name="wapProductInfo" type="text" id="wapProductInfo" value="<%=rs("wapproductinfo")%>" size="3"></td>
        </tr>
        <tr>
          <td >WAP图片列表页调用条数：</td>
		   <td >
            <input name="wapalinfo" type="text" id="wapalinfo" value="<%=rs("wapalinfo")%>" size="3"></td>
        </tr>
        <tr>
          <td >WAP下载列表页调用条数：</td>
		   <td >
            <input name="wapdowninfo" type="text" id="wapdowninfo" value="<%=rs("wapdowninfo")%>" size="3"></td>
        </tr>
		
        <tr>
          <td >WAP视频列表页调用条数：</td>
		   <td >
            <input name="waptvinfo" type="text" id="waptvinfo" value="<%=rs("waptvinfo")%>" size="3"></td>
        </tr>
		
        <tr>
          <td >WAP招聘列表页调用条数：</td>
		   <td >
          <input name="wapzpinfo" type="text" id="wapzpinfo" value="<%=rs("wapzpinfo")%>" size="3"></td>
        </tr>
		<tr>
		<tr>
          <td >WAP留言列表页显示条数：</td>
		   <td >
            <input name="waplyinfo" type="text" id="waplyinfo" value="<%=rs("waplyinfo")%>" size="3"></td>
        </tr>
		<%end if%>
		
		<%if t=3 then%>
  <tr>
          <td width="18%">地图坐标： </td>
		   <td ><a href="http://api.map.baidu.com/lbsapi/getpoint/index.html" target="_blank" ><font color=red>点此获取</font></a>
      <input name="wapmap" type="text" id="wapmap" value="<%=rs("wapmap")%>" size="30">
      (比如：111.45824,35.481948)</td>
    </tr>
  <tr>
          <td >公司名称：</td>
		   <td >
          <input name="wapmc" type="text" id="wapmap" value="<%=rs("wapmc")%>" size="30"></td>
    </tr>
	 <tr>
          <td >联系电话： </td>
		   <td >
      <input name="mapphone" type="text" id="mapphone" value="<%=rs("mapphone")%>" size="30">
      </td>
    </tr>

 <script>
 var editor;
 KindEditor.ready(function(K) {
 editor = K.create('textarea[name="wapcontent"]', {
 filterMode : false, 
 uploadJson : '../zmbjq/asp/upload_json.asp',
 fileManagerJson : '../zmbjq/asp/file_manager_json.asp',
 allowFileManager : true,
 width: '700px',
  hight: '300px',
 afterBlur: function(){this.sync();}
 });

 });
 </script>
     <tr>
    <td >备注：</td>
		   <td >
      <label>
      <textarea name="wapcontent" id="wapcontent"  style="width:600px;height:300px;visibility:hidden;"><%=rs("wapcontent")%></textarea>
      </label></td>
  </tr>
  <%end if%>
   <%if t=4 then%>
      <tr >
  <td width="18%">手机LOGO图片：</td>
		   <td >
				   <input name="ewm"  type="text" id="weepic"  value="<%=rs("ewm")%>" size="50"  /><br/><br/>
			      <a class='sfile' href="Upload_All.Asp?add=one" title="支持上传格式(*.swf;*.gif;*.bmp)"><div style="width:100px;padding:10px; background:#28b576;color:#ffffff"">二维码上传</div></a>
		        <div id="mysmaimg"><%if rs("ewm")<>"" then%>
<img src="../<%=rs("ewm")%>" id="sc"/><br />
<a href="delslt.asp?sltdz=<%=rs("ewm")%>" target="_blank" onClick="deltp()" id="lj">删除</a>
<%end if%></div>		       </td>
    </tr>
		
<tr> <td>手机LOGO大小：</td>
		   <td >
	  
          <input name="ewmwidth"type="text" class="inputtext" id="ewmwidth"  value="<%=rs("ewmwidth")%>" size="6" maxlength="30" />
                    宽×高
        <input name="ewmheight"type="text" class="inputtext" id="ewmheight"  value="<%=rs("ewmheight")%>" size="6" maxlength="30" /></td>
    </tr>
	<%end if%>
	
	 <%if t=5 then%>
      <tr >
  <td  width="18%">微信二维码图片：</td>
		   <td >
				   <input name="wxewm"  type="text" id="weepic"  value="<%=rs("wxewm")%>" size="50"  /><br/><br/>
			      <a class='sfile' href="Upload_All.Asp?add=one" title="支持上传格式(*.swf;*.gif;*.bmp)"><div style="width:100px;padding:10px; background:#28b576;color:#ffffff" align="center">微信二维码上传</div></a>
		        <div id="mysmaimg"><%if rs("wxewm")<>"" then%>
<img src="../<%=rs("wxewm")%>" id="sc"/><br />
<a href="delslt.asp?sltdz=<%=rs("wxewm")%>" target="_blank" onClick="deltp()" id="lj">删除</a>
<%end if%></div>		        </td>
    </tr>
		
<tr> <td>微信二维码大小：</td>
		   <td >
	  
          <input name="wxewmwidth"type="text" class="inputtext" id="wxewmwidth"  value="<%=rs("wxewmwidth")%>" size="6" maxlength="30" />
                    微信二维码宽×高
        <input name="wxewmheight"type="text" class="inputtext" id="wxewmheight"  value="<%=rs("wxewmheight")%>" size="6" maxlength="30" /></td>
    </tr>
	<%end if%>
	
	
	
	
	
	
	
      <tr>
    <td colspan="2">
      <input type="submit" name="Submit" value="保存" class="bnt" />
    </td>
  </tr>
   
</table>
</form>

</body>
</html>
<%
rs.nothing
set rs=nothing
 conn.Close
    Set conn = Nothing
%>
