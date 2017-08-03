<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!-- #include file="admin.asp" -->
<!-- #include file="Inc/function.asp" -->
<!-- #include file="Inc/Head.asp" -->
<!--#include file="sc.asp"-->
<%
theInstalledObjects = "Persits.Jpeg"
if Instr(session("AdminPurview"),"|91,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</script>"
  response.end
end if
'========判断是否具有管理权限
%>
<%
If request("action") = "del" Then


sytp=request.QueryString("sytp")
call deletefile(sytp)
Set rs = server.CreateObject("adodb.recordset")
sql = "select * from wzset  "
rs.Open sql, conn, 1, 3
rs("sytp")=null
rs.update
    rs.Close
    Set rs = Nothing
   
  
	call addlog("删除水印图片")


end if
%>
<%  


PicW = request("PicW")
PicH = request("PicH")
PicBW=request("PicBW")
PicBH=request("PicBH")
Pictxt=request("Pictxt")
picLocation=request("picLocation")
PicColor=request("PicColor")
PicSize=request("PicSize")
Piclx=request("Piclx")
sytp=request("sytp")

If request.Form("submit") = "保存" Then

Set rs = server.CreateObject("adodb.recordset")
sql = "select * from wzset where id=1 "
rs.Open sql, conn, 1, 3
 rs("PicW")= PicW
 rs("PicH")= PicH
rs("PicBW")=PicBW
rs("PicBH")=PicBH
rs("Pictxt")=Pictxt
rs("picLocation")=picLocation
rs("PicColor")=PicColor
rs("PicSize")=PicSize
rs("Piclx")=Piclx
rs("sytp")=sytp
rs.update
    rs.Close
    Set rs = Nothing
   
	call addlog("设置缩略图")
  response.Write "<script language=javascript>alert(""恭喜您修改成功"");window.location.href=""slt.asp""</script>"
    response.End
End If
%>


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>zmcms_管理页面</title>
<link rel="stylesheet" href="css/jquery.bigcolorpicker.css" type="text/css" />
<script type="text/javascript" src="js/jquery-1.7.2.min.js"></script>

	<style type="text/css">
table{ border:1px solid #D4D4D4; font-family: Arial, Helvetica, sans-serif;font-size: 12px; background-color:#faf9d1}
table td{BORDER-bottom: #D4D4D4 1px dashed;line-height:30px; padding:5px 0 10px 10px}
</style>

</head>


<body>

<form method="POST" action="slt.asp" id="form" name="form">
<%
Set rs = server.CreateObject("adodb.recordset")
sql = "select * from wzset where id=1"
rs.Open sql, conn, 1, 1
If rs.EOF Then
response.Write "<script language=javascript>alert(""不存在此数据！"");window.history.back();</script>"
response.End
End If
%>

<table width="99%" border="0" align="center" cellpadding="0" cellspacing="0" style="border-bottom:none; background:#eeeeee;">
  <tr>
    <td style="padding:2px; border-bottom:none; border-right:#dddddd solid 1px"><div align="center"><a href="wzset.asp"target="main">基本设置</a></div></td>
    <td style="padding:2px; border-bottom:none; border-right:#dddddd solid 1px"><div align="center"><a href="lbsm.asp" target="main">列表页数量</a></div></td>
    <td style="padding:2px; border-bottom:none; border-right:#dddddd solid 1px"><div align="center"><a href="cj.asp" target="main">插件是否开启</a></div></td>
    <td style="padding:2px; border-bottom:none; border-right:#dddddd solid 1px"><div align="center"><a href="qtsz.asp" target="main">其他设置</a></div></td>
    <td style="padding:2px; border-bottom:none; border-right:#dddddd solid 1px"><div align="center"><a href="slt.asp" target="main">缩略图参数</a>&nbsp;<font color="#FF0000">new</font></div></td>
    <td style="padding:2px; border-bottom:none; border-right:#dddddd solid 1px"><div align="center"><a href="emaill.asp" target="main">邮箱配置</a>&nbsp;<font color="#FF0000">new</font></div></td>
    <td style="padding:2px; border-bottom:none; border-right:#dddddd solid 1px"><div align="center"><a href="share.asp"target="main">分享参数</a></div></td>
    <td style="padding:2px; border-bottom:none; border-right:#dddddd solid 1px"><div align="center"><a href="kfcs.asp"target="main">在线客服参数</a></div></td>
    <td style="padding:2px; border-bottom:none; border-right:#dddddd solid 1px"><div align="center"><a href="messageskin.asp"target="main">客户即时咨询样式</a>&nbsp;<font color="#FF0000">new</font></div>
      <div align="center"></div>
      <div align="center"></div>
      <div align="center"></div>
      <div align="center"></div>
      <div align="center"></div>      <div align="center"></div></td>
    </tr>
</table>
<table width="99%"  border="0" cellpadding="6" cellspacing="0"  bgcolor="#f9f9f9" align="center" >
                                      
                                      

<tr >
  <td height="25" align="left" >小图设置（宽度）：
                    
                      <input name="PicW" type="text" id="tpl" value="<%=rs("PicW")%>" size="10"   />                 </td>
</tr>
<tr >
                                        <td height="25" align="left" >小图设置（高度）：
                                          <input name="PicH" type="text"  value="<%=rs("PicH")%>" size="10" />                   </td>
                                      </tr>
                                      

<tr >
  <td height="25" >大图设置（宽度）：
                      <input name="PicBW" type="text"  value="<%=rs("PicBW")%>" size="10" />  </td>
</tr>
<tr >
  <td height="25" >大图设置（高度）：<input name="PicBH" type="text"  value="<%=rs("PicBH")%>" size="10" /></td>
</tr>
<tr >
  <td height="25" >文字水印内容：<input name="Pictxt" type="text" value="<%=rs("Pictxt")%>" size="15" /></td>
</tr>
<tr >
  <td height="25" >文字水印位置：<select name="picLocation" >
                      <option value="0"  <%if rs("picLocation")=0 then Response.Write("selected")%>>左上</option>
                      <option value="1"  <%if rs("picLocation")=1 then Response.Write("selected")%>>右上</option>
                      <option value="2"  <%if rs("picLocation")=2 then Response.Write("selected")%>>左下</option>
                      <option value="3"  <%if rs("picLocation")=3 then Response.Write("selected")%>>右下</option>
                      <option value="4">居中</option>
                    </select>    </td>
</tr>
<tr >
  <td height="25" >文字水印颜色：
    <input name="PicColor" type="text"  value="<%=rs("PicColor")%>" size="15" id="PicColor"  title="文字水印颜色,请只输入6位颜色代码,如白色FFFFFF" /></td>
</tr>
<tr >
  <td height="25" >文字水印大小：<input name="PicSize" type="text"  value="<%=rs("PicSize")%>" size="15" id="tpl"  title="水印字体大小,请只需要输入数字,以像素(PX)为默认单位" />    </td>
</tr>

<tr >
  <td height="25" >是否为图片水印：<select name="Piclx" >
     <option value="0"  <%if rs("Piclx")=0 then Response.Write("selected")%>>关闭水印缩略图功能</option>
     <option value="1"  <%if rs("Piclx")=1 then Response.Write("selected")%> <%If Not IsObjInstalled(theInstalledObjects) Then Response.Write("disabled") %>>使用文字水印+缩略图功能</option>
     <option value="2"  <%if rs("Piclx")=2 then Response.Write("selected")%>  <%If Not IsObjInstalled(theInstalledObjects) Then Response.Write("disabled") %>>使用图片水印+缩略图功能</option>
     <option value="3"  <%if rs("Piclx")=3 then Response.Write("selected")%>  <%If Not IsObjInstalled(theInstalledObjects) Then Response.Write("disabled") %>>仅使用缩略图功能</option>
                                          </select></td>
</tr>
<tr >
  <td height="25" >水印图片：<input name="sytp" type="text" id="url3" size="30" value="<%=rs("sytp")%>" />
<input type="button" id="image3" value="上传图片" />  <br /><% if not IsFileExist(server.mappath(".."&sitepath&""&rs("sytp")&"")) then 
		response.Write("<font color=red>请上传您的水印图片</font>")   
		%>
			
		<%
		else
		%>
		<img src="../<%=rs("sytp")%>"><br>
		<a href="?action=del&sytp=<%=rs("sytp")%>"><div style="width:80px;border:1px #ddd solid; background:#3399FF; text-align:center; color:#fff">删除</div></a>  注：上传新水印图片之前，请先删除以前的水印图片
		<%
		end if
		
		%>	</td>
</tr>
<tr >
                                        <td height="25" >
  <input  class="bnt"type="submit" value="保存" id="submit4" name="submit" />                                        </td>
                                      </tr>
   </table>
</form>

</body>
</html>
<%
 conn.Close
    Set conn = Nothing
%>
