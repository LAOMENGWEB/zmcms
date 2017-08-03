<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!-- #include file="admin.asp" -->
<!-- #include file="Inc/function.asp" -->
<%
if Instr(session("AdminPurview"),"|91,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</script>"
  response.end
end if
'========判断是否具有管理权限
If request.Form("submit") = "保存" Then
wzkg = request("wzkg")
wh = request("wh")
mc=request("mc")
logo=request("logo")
bq=request("bq")
logowidth=request("logowidth")
logoheight=request("logoheight")
ms=request("ms")
gjc=request("gjc")
ym=request("ym")
yx=request("yx")
dw=request("dw")
map=request("map")
lxdz=request("lxdz")
lxr=request("lxr")
gddh=request("gddh")
sj=request("sj")
ba=request("ba")
cz=request("cz")
tvid=request("tvid")
dcid=request("dcid")
qyid=request("qyid")
zid=request("zid")
ljid=request("ljid")
Set rs = server.CreateObject("adodb.recordset")
sql = "select * from wzset where lang='"&lang&"' "
rs.Open sql, conn, 1, 3
rs("wzkg")= wzkg 
rs("wh")= wh
rs("mc")=mc
rs("ms")=ms
rs("gjc")=gjc
rs("ym")=ym
rs("yx")=yx
rs("dw")=dw
rs("lxdz")=lxdz
rs("lxr")=lxr
rs("gddh")=gddh
rs("sj")=sj
rs("ba")=ba
rs("cz")=cz
rs("bq")=bq
rs("map")=map
rs("lang")=lang
rs("tvid")=tvid
rs("dcid")=dcid
rs("qyid")=qyid
rs("zid")=zid
rs("ljid")=ljid
rs.update
    rs.Close
    Set rs = Nothing
   
	call addlog("修改基本设置")
  response.Write "<script language=javascript>alert(""恭喜您修改成功"");window.location.href=""wzset.asp""</script>"
    response.End
End If
%>


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>zmcms_管理页面</title>

<script language="javascript" src="../js/Admin.js"></script>
<script type="text/javascript" src="js/jquery-1.7.2.min.js"></script>
<script type="text/javascript" src="../js/Validform_v5.3.2_min.js"></script>
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

<link href="css/style.css" rel="stylesheet" type="text/css">
 <script type="text/javascript">
$(function(){
	
$(".checkform").Validform();  //就这一行代码！;
	
})
</script>
<style type="text/css">
table{ font-size: 14px;  border-right:none}
table td{padding:10px; font-family:"微软雅黑";color:#ccc; font-weight:bold;}
table td a{color:#ffffff}
</style>
<!--#include file="hqyy.asp"-->
</head>


<body>

<form name="editForm" method="POST" class="checkform"  >
<%
Set rs = server.CreateObject("adodb.recordset")
sql = "select * from wzset where lang='"&lang&"'"
rs.Open sql, conn, 1, 1
%>		
<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" style=" background: #F9f9f9;border:none; ">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">基本设置  <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>

<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="set.asp" target="main">返回菜单</a></div></td>
 

 </tr>
</table>

<table width="99%" border="0" align="center" cellpadding="0" cellspacing="0"  class="dtable4">
  <tr>
    <td >网站是否关闭：	</td>
    <td colspan="2">
	<input style="border:none" type="radio" name="wzkg" value="1" <%if rs("wzkg")="1" then response.write "checked"%> >
关闭
<input style="border:none" type="radio" name="wzkg" value="2" <%if rs("wzkg")="2" then response.write "checked"%>>
不关闭</td>
    </tr>
  <script>
 var editor;
 KindEditor.ready(function(K) {
 editor = K.create('textarea[name="wh"]', {
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
	 <td>网站关闭显示内容：</td>
    <td colspan="2" >
     <textarea name="wh" id="wh" style="width:400px;height:200px;visibility:hidden;"><%=rs("wh")%></textarea>      </td>
  </tr>
   
 
  
      <tr>
	  <td ><span style="color:#ff0000">*</span> 网站名称：</td>
        <td colspan="2">
       
        <input name="mc" type="text" id="mc2" size="80" value="<%=rs("mc")%>" datatype="*" nullmsg="网站名称不能为空"/>               </td>
      </tr>
	  
	  <tr>
        <td ><span style="color:#ff0000">*</span> 网站网址：</td>
        <td colspan="2"><input name="ym" type="text" id="mc" size="80" value="<%=rs("ym")%>" datatype="*" nullmsg="网站网址不能为空"/>   例如: http://www.zmxt.cn</td>
      </tr>


		

      <tr>
	  <td><span style="color:#ff0000">*</span> 站点描述： </td>
        <td colspan="2">
        <textarea datatype="*" nullmsg="站点描述不能为空" name="ms" cols="50" rows="5"  style="border:1px #ddd solid"><%=rs("ms")%></textarea>      </td>
      </tr>
      <tr>
	  <td><span style="color:#ff0000">*</span> 站点关键词：</td>
        <td>
          <textarea datatype="*" nullmsg="站点关键词不能为空" name="gjc" cols="50" rows="5" style="border:1px #ddd solid"><%=rs("gjc")%></textarea>        </td>
      </tr>
      <tr>
	  <td>网站邮箱：</td>
        <td colspan="2">
      <input name="yx" type="text" id="yx" value="<%=rs("yx")%>" size="40" maxlength="50">   </td>
      </tr>
     <tr>
	 <td>单位名称：</td>
      <td colspan="2" >
      <input name="dw" type="text" id="dw" value="<%=rs("dw")%>" size="40" maxlength="255">  </td>
    </tr>
    <tr>
	<td>单位地址：</td>
      <td colspan="2" >
      <input name="lxdz" type="text" id="lxdz" value="<%=rs("lxdz")%>" size="40" maxlength="255"> </td>
    </tr>
    <tr>
	<td>联系人：</td>
      <td colspan="2" >
      <input name="lxr" type="text" id="lxr" value="<%=rs("lxr")%>" size="40" maxlength="255">      </td>
    </tr>
	<tr>
	<td>电话号码：</td>
      <td colspan="2" >
      <input name="gddh" type="text" id="gddh" value="<%=rs("gddh")%>" size="40" maxlength="255">      </td>
    </tr>

   <tr>
   <td>手机号码：</td>
      <td colspan="2" >
      <input name="sj" type="text" id="sj" value="<%=rs("sj")%>" size="40" maxlength="255">      </td>
    </tr>
	<tr>
	<td>备案号：</td>
      <td colspan="2" >
      <input name="ba" type="text" id="ba" value="<%=rs("ba")%>" size="40" maxlength="255">     </td>
    </tr>
	<tr>
	<td>公司传真：</td>
      <td colspan="2" >
      <input name="cz" type="text" id="cz" value="<%=rs("cz")%>" size="40" maxlength="255"></td>
    </tr>
	
    
    <tr >
	<td>版权信息：</td>
      <td colspan="2" >
       
        <input name="bq" type="text" id="bq" value="<%=rs("bq")%>" size="100" />     </td>
    </tr>
     <tr >
	<td>首页视频ID调用：</td>
      <td colspan="2" >
       
        <input name="tvid" type="text" id="bq" value="<%=rs("tvid")%>" size="100" />     </td>
    </tr>
     <tr >
	<td>首页调查ID调用：</td>
      <td colspan="2" >
       
        <input name="dcid" type="text" id="bq" value="<%=rs("dcid")%>" size="100" />     </td>
    </tr>
         <tr >
	<td>首页企业简介ID调用：</td>
      <td colspan="2" >
       
        <input name="qyid" type="text" id="bq" value="<%=rs("qyid")%>" size="100" />     </td>
    </tr>
          <tr >
	<td>指定单页分类ID：</td>
      <td colspan="2" >
       
        <input name="zid" type="text" id="bq" value="<%=rs("zid")%>" size="100" />     </td>
    </tr>
    
            <tr >
	<td>指定链接分类ID：</td>
      <td colspan="2" >
       
        <input name="ljid" type="text" id="bq" value="<%=rs("ljid")%>" size="100" />     </td>
    </tr>
    <tr >
	<td>地图代码：</td>
      <td colspan="2" >
	  <a href="http://api.map.baidu.com/lbsapi/creatmap/index.html"> 百度地图申请</a><br /><br />
	  
	  <textarea name="map" cols="100" rows="30" id="map"><%=rs("map")%></textarea></td>
    </tr>

      <tr>
	  
    <td colspan="2">
      <input type="submit" name="Submit" value="保存" class="bnt" />   </td>
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
