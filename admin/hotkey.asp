<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="Inc/Function.asp"-->
<!--#include file="class_page.asp"-->
<%
Result=request.QueryString("Result")
keytag=request.Form("k1")
pone=request.QueryString("pone")
if Instr(session("AdminPurview"),"|10414,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</script>"
  response.end
end if
'========判断是否具有管理权限
%>
<% 
If request.Form("del") = "批量删除" Then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='hotkey.asp';</script>" 
response.end()
end if
conn.Execute ( "delete from hotkey where id in ("&Request("id")&")" )
response.Write "<script language=javascript>alert(""删除成功！"");window.location.href=""hotkey.asp""</script>"
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>zmcms_管理页面</title>

<script language="JavaScript" src="js/jquery-1.7.2.min.js"></script>
 <script type="text/javascript" src="../js/Validform_v5.3.2_min.js"></script>
 <link href="css/style.css" rel="stylesheet" type="text/css">
  <script type="text/javascript">
$(function(){
	
$(".checkform").Validform();  //就这一行代码！;
	
})
</script>
		<SCRIPT language=javascript>
function CheckAll(){ 
 for (var i=0;i<eval(zmqy.elements.length);i++){ 
  var e=zmqy.elements[i]; 
  if (e.name!="allbox") e.checked=zmqy.allbox.checked; 
 } 
} 
function ConfirmDel()
{
   if(confirm("确定要删除选中的热搜关键词吗？一旦删除将不能恢复！"))
     return true;
   else
     return false;
	 
}

</SCRIPT>
<style type="text/css">
.table1{ font-size: 14px; background: #F9f9f9;border:none;}
.table1 td{padding:10px; font-family:"微软雅黑";color:#ccc; font-weight:bold;}
.table1 td a{color:#ffffff}
</style>
<!--#include file="hqyy.asp"-->
</head>
<body>
<form id="form2" name="form2" method="post" action="hotkey.asp?Result=ok" >


<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">热门搜索词管理 <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>


<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="chajian.asp" target="main">返回菜单</a></div></td>
 

 </tr>
</table>

<table width="99%" align="center" cellpadding="0" cellspacing="0" class="dtable1">
  		
          <tr   >
            <td width="9%" >
             快速搜索：</td>
			 <td width="91%">
                <input name="k1" type="text" id="k1" />
             
                            <label><input class="bnt" type="submit" name="button" id="button" value="搜索" /></label>
            </td>
          </tr>
  </table>
  </form>
<form name="zmqy" method="Post" >
 <%
 if Result="ok"   then
 if keytag="" then
 Response.Write "<script>alert('搜索关键字不能为空');window.location.href='hotkey.asp';</script>" 
response.end()
 end if
 DB_Sql = "select * from hotkey where keywords like '%"&keytag&"%' and lang='"&lang&"' "
else
DB_Sql = "select * from hotkey where lang='"&lang&"' "
end if
DB_Sql = DB_Sql & "order by ID desc"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open DB_Sql,Conn,1,1
If Not(rs.Bof and rs.Eof) Then


'================================================
'调用分页类开始
Dim Obj_Page, Obj_PageListnum, II
Obj_PageListnum = 20 

Set Obj_Page = New HHYY_Pageclass

Obj_Page.Setpage_RS = rs
Obj_Page.Setpage_Listnum = Obj_PageListnum

Obj_Page.Setpage_Pathurl = "hotkey.asp" '执行分页的页面
Obj_Page.Setpage_Navtype = 1 '分页导航样式(可选1文字,2数字)

Call Obj_Page.Page_Init() '初始化分页
%>
<table width="99%"   align="center" cellpadding="0" cellspacing="0" class="dtable2">
<tr  class="wz">
	<td width="3%"  align="center"  >全选</td>
    <td width="9%" align="center"  >编号</td>
    <td width="11%"  align="center"  >关键字名称</td>
    <td width="13%" align="center"  >搜索次数</td>
    <td width="13%" align="center"  >搜索关键词所属分类</td>
    <td width="13%"  align="center"  >最近搜索时间</td>
    <td width="20%" align="center"  >搜索用户IP</td>
    <td width="18%"  align="center"  style="border-right:none">管理</td>    
  </tr>
<%
'/循环开始
For II = 1 to Obj_PageListnum
If rs.Eof Then
	Exit For
End If
%>
    <tr >
    <td ><div align="center">
  <input name="ID" type="checkbox" id="ID" value="<%=rs("id")%>" style="border:none"/>
 </div>      </td>
    <td><div align="center"><%=rs("id")%></div></td>
    <td height="25"><div align="center"><%=rs("keywords")%></div></td>
    <td height="25" align="center"><div align="center">
     <iframe name="zmcms" width="150" height="50" frameborder="0" scrolling="no" 
				src="xghotkey.asp?id=<%=rs("id")%>" ></iframe>    </div></td>
    <td align="center">
	  <div align="center">
	    <%
	if rs("ssfl")=1 then
 response.Write("新闻")
end if
	if rs("ssfl")=2 then
 response.Write("产品")
end if
	if rs("ssfl")=3 then
 response.Write("图片")
end if
	if rs("ssfl")=4 then
 response.Write("视频")
end if
	if rs("ssfl")=5 then
 response.Write("下载")
end if

%>
      </div></td>
    <td height="25" align="center"><div align="center"><%=rs("adddate")%></div></td>
    <td align="center"><%=rs("sip")%></td>
    <td align="center" style="border-right:none">
	<a href="delhotkey.asp?id=<%=rs("ID")%>" onClick="JavaScript:return confirm('确定要删除吗？')">删除</a>	</td>    
  </tr>

 <%
rs.MoveNext
Next	
%> 

	<tr>
    <td height="25" align="center" ><input type="checkbox" name="allbox" onClick="CheckAll()" style="border:none"/></td>
    <td height="25" colspan="7" style="padding:5px ; text-align:left"><input name="Del" type="submit" class="bnt" id="Del"  onClick="JavaScript:return confirm('删除吗？')" value="批量删除"></td>
    </tr>
</table>
<%If Obj_Page.OutErrInfo <> "" Then
		Response.Write Obj_Page.OutErrInfo '报错信息显示
		Response.End
		else
	Call Obj_Page.Page_Navigation() '分页导航
	end if
    else
	%>
<table width="99%" border="0" align="center" cellpadding="0" cellspacing="0" class="table1" style="margin-top:10px">
  <tr>
    <td style="padding:10px 0 ;border-right:none;border-bottom:none">	<div align="center">
      <%
Response.Write "目前暂无任何信息"
Response.end
	%>
    </div></td>
  </tr>
</table>

<%Set Obj_Page = Nothing

RS.Close
Set RS = Nothing
End If
%>
</form>
</body>
</html>
