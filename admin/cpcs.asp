<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="inc/function.asp"-->
<%
if Instr(session("AdminPurview"),"|35,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</script>"
  response.end
end if
'========判断是否具有管理权限
%>
<% 
If request.Form("submit") = "确认修改" Then
skins=request("skins")
timess=Request("timess")
triggers=Request("triggers")
txtheights=Request("txtheights")
wraps=Request("wraps")
slt=Request("slt")
sltgd=Request("sltgd")
Set rs = server.CreateObject("adodb.recordset")
sql = "select * from cp "
rs.Open sql, conn, 1, 3
rs("skins") = skins
rs("triggers") =triggers
rs("timess") =timess
rs("txtheights") =txtheights
rs("wraps") =wraps
rs("slt")=slt
rs("sltgd")=sltgd
    rs.update
    rs.Close
    Set rs = Nothing
    conn.Close
    Set conn = Nothing
    response.Write "<script language=javascript>alert(""恭喜您修改成功！"");window.location.href=""cpcs.asp""</script>"
    response.End
End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>zmcms_管理页面</title>
<script language="JavaScript" src="../js/jquery-1.7.2.min.js"></script>
 <script type="text/javascript" src="../js/Validform_v5.3.2_min.js"></script>
 <link href="css/style.css" rel="stylesheet" type="text/css">
  <script type="text/javascript">
$(function(){
	
$(".checkform").Validform();  //就这一行代码！;
	
})
</script>
</head>

<body>
<form  name="mx" method="post" class="checkform">
<%
Set rs = server.CreateObject("adodb.recordset")
sql = "select * from cp"
rs.Open sql, conn, 1, 1
If rs.EOF Then
response.Write "<script language=javascript>alert(""不存在此数据！"");window.history.back();</script>"
response.End
End If
%>


  <table width="100%" border="0" cellpadding="0" cellspacing="0"   class="dtable">
<tr>
 <td ><%=chanpin%>幻灯参数设置</td>
</tr>
</table>

<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="chanpin.asp" target="main">返回菜单</a></div></td>
 

 </tr>
</table>
                      
                         <table width="99%"  border="0" cellpadding="0" cellspacing="0"  align="center" class="dtable4">
                                   
                                      <tr>
                                        <td width="13%"  >产品显示样式：</td>
										
										<td width="87%">
                                          <select 
name="skins" id="skins">
 
 <option value="cpys2" <%if rs("skins")="cpys2" then Response.Write("selected")%>>新增样式</option>


                                            <option value="mF_51xflash" <%if rs("skins")="mF_51xflash" then Response.Write("selected")%>>幻灯样式一</option>
                                            <option value="mF_classicHB" <%if rs("skins")="mF_classicHB" then Response.Write("selected")%>>幻灯样式二</option>
                                            <option value="mF_classicHC" <%if rs("skins")="mF_classicHC" then Response.Write("selected")%>>幻灯样式三</option>
                                            <option value="mF_expo2010" <%if rs("skins")="mF_expo2010" then Response.Write("selected")%>>幻灯样式四</option>
                                            <option value="mF_fscreen_tb" <%if rs("skins")="mF_fscreen_tb" then Response.Write("selected")%>>幻灯样式五</option>
                                            <option value="mF_games_tb" <%if rs("skins")="mF_games_tb" then Response.Write("selected")%>>幻灯样式六</option>
                                            <option value="mF_kiki" <%if rs("skins")="mF_kiki" then Response.Write("selected")%>>幻灯样式七</option>
                                            <option value="mF_ladyQ" <%if rs("skins")="mF_ladyQ" then Response.Write("selected")%>>幻灯样式八</option>
                                            <option value="mF_liquid" <%if rs("skins")="mF_liquid" then Response.Write("selected")%>>幻灯样式九</option>
                                            <option value="mF_liuzg" <%if rs("skins")="mF_liuzg" then Response.Write("selected")%>>幻灯样式十</option>
                                            <option value="mF_luluJQ" <%if rs("skins")="mF_luluJQ" then Response.Write("selected")%>>幻灯样式十一</option>
                                            <option value="mF_pconline" <%if rs("skins")="mF_pconline" then Response.Write("selected")%>>幻灯样式十二</option>
                                            <option value="mF_peijianmall" <%if rs("skins")="mF_peijianmall" then Response.Write("selected")%>>幻灯样式十三</option>
                                            <option value="mF_pithy_tb" <%if rs("skins")="mF_pithy_tb" then Response.Write("selected")%>>幻灯样式十四</option>
                                            <option value="mF_qiyi" <%if rs("skins")="mF_qiyi" then Response.Write("selected")%>>幻灯样式十五</option>
                                            <option value="mF_quwan" <%if rs("skins")="mF_quwan" then Response.Write("selected")%>>幻灯样式十六</option>
                                            <option value="mF_rapoo" <%if rs("skins")="mF_rapoo" then Response.Write("selected")%>>幻灯样式十七</option>
                                            <option value="mF_slide3D" <%if rs("skins")="mF_slide3D" then Response.Write("selected")%>>幻灯样式十八</option>
                                            <option value="mF_sohusports" <%if rs("skins")="mF_sohusports" then Response.Write("selected")%>>幻灯样式十九</option>
                                            <option value="mF_taobao2010" <%if rs("skins")="mF_taobao2010" then Response.Write("selected")%>>幻灯样式二十</option>
                                            <option value="mF_taobaomall" <%if rs("skins")="mF_taobaomall" then Response.Write("selected")%>>幻灯样式二十一</option>
                                            <option value="mF_tbhuabao" <%if rs("skins")="mF_tbhuabao" then Response.Write("selected")%>>幻灯样式二十二</option>
                                            <option value="mF_YSlider" <%if rs("skins")="mF_YSlider" then Response.Write("selected")%>>幻灯样式二十三</option>
											<option value="mF_dleung" <%if rs("skins")="mF_dleung" then Response.Write("selected")%>>幻灯样式二十四</option>
											<option value="mF_fancy" <%if rs("skins")="mF_fancy" then Response.Write("selected")%>>幻灯样式二十五</option>
											<option value="mF_kdui" <%if rs("skins")="mF_kdui" then Response.Write("selected")%>>幻灯样式二十六</option>
											<option value="mF_shutters" <%if rs("skins")="mF_shutters" then Response.Write("selected")%>>幻灯样式二十七</option>
                                        </select>
                                   </td>
                                      </tr>
                                      
                                      
                                      <tr>
                                        <td  >效果切换时间：</td>
										
										<td>
                                        <input name="timess" type="text" id="timess"  value="<%=rs("timess")%>" size="2" /></td>
                                      </tr>
                                      <tr>
                                        <td  >文字层高度设置：</td>
										
										<td>
                                          <input name="txtHeights" type="text" id="txtHeights"  value="<%=rs("txtHeights")%>" size="2" /> 
                                          ( 0为隐藏)</td>
                                      </tr>
                                      <tr>
                                        <td  >是否显示边框：</td>
										
										<td>
                                          <select name="wraps" id="wraps" >
                                            <option value="true"  <%if rs("wraps")="true" then Response.Write("selected")%>>显示边框</option>
                                            <option value="false" <%if rs("wraps")="false" then Response.Write("selected")%>>关闭边框</option>
                                          </select></td>
                                      </tr>
                                      <tr>
                                        <td  >触发模式：</td>
										
										<td>
                                          <select name="triggers" id="triggers"  >
                                            <option value="click"  <%if rs("trigger")="click" then Response.Write("selected")%>>鼠标点击</option>
                                            <option value="mouseover"  <%if rs("trigger")="mouseover" then Response.Write("selected")%>>鼠标悬停</option>
                                        </select>
                                     </td>
                                      </tr>
									      <tr >
                                        <td   >缩略图个数：</td>
										
										<td>
                                        
                                          <input name="slt" type="text" id="slt" value="<%=rs("slt")%>" size="8" />
                                        （ 缩略图个数和高度 这2个参数只针对于 幻灯参数5和6）</td>
                                      </tr>
                                      <tr >
                                        <td   >缩略图高度：</td>
										
										<td>
                                          <label>
                                          <input name="sltgd" type="text" id="sltgd" value="<%=rs("sltgd")%>" size="8" />
                                        </label></td>
                                      </tr>
                                     <tr>
                                       <td  colspan="2" >
  <input type="submit" value="确认修改" id="submit4" name="submit" class="bnt"/></td>
                           </tr> 
   </table>
   </form>
</body>
</html>
