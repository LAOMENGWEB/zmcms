<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="class_page.asp"-->
<!--#include file="Inc/Function.asp"-->
<%
if Instr(session("AdminPurview"),"|511,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</script>"
  response.end
end if
'========判断是否具有管理权限
%>
<%
If request.Form("button") = "批量删除" Then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='psfs.asp';</script>" 
response.end()
end if
dim id
ID=trim(request("ID"))
ID1=split(ID,",")
for i=0 to ubound(ID1)
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from shoppingshipment where ID in ("&id1(i) &")",conn,0,1
dim picUrl
dim picUrlArray
dim x,y
picUrl = getPicUrl(rs("content"))
if picUrl <> "" then
picUrl = left(picUrl,len(picUrl)-1)
picUrlArray = split(picUrl,",")
for x = 0 to ubound(picUrlArray)
if instr(picUrlArray(x),"") > 0 then
call DeleteFile(picUrlArray(x))
end if
next
end if
next
conn.Execute ( "delete from shoppingshipment where id in ("&Request("id")&")" )
call addlog("删除配送方式")
response.Write "<script language=javascript>alert(""删除成功！"");window.location.href=""psfs.asp""</script>"
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
   if(confirm("确定要删除选中的配送方式吗？一旦删除将不能恢复！"))
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
<form name="zmqy" method="Post">

<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">配送方式管理  <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>



<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="chajian.asp" target="main">返回菜单</a></div><div class="dmeun"><a href="psfsadd.asp" target="main">添加配送方式</a></div></td>
 

 </tr>
</table>
<%
DB_Sql = "select * from shoppingshipment where lang='"&lang&"' "
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

Obj_Page.Setpage_Pathurl = "psfs.asp" '执行分页的页面
Obj_Page.Setpage_Navtype = 1 '分页导航样式(可选1文字,2数字)

Call Obj_Page.Page_Init() '初始化分页
%>

              <table width="99%" align="center" cellpadding="0" cellspacing="0" class="dtable2">
                
                <tr  class="wz">
				<td width="13%" > 
                  <div align="center">全选</div></td>
                  <td width="16%"   ><div align="center">ID</div></td> 
                  <td width="23%"  > 
                  <div align="center">配送名称</div></td>
                  <td width="25%"  ><div align="center">配送金额</div></td>
                  <td width="23%"  style="border-right:none"><div align="center">操作</div></td>
                </tr>
<%
'/循环开始
For II = 1 to Obj_PageListnum
If rs.Eof Then
	Exit For
End If
%>
               
                <tr  >
				 <td><div align="center">
                   <input name="ID" type="checkbox" id="ID" value="<%=rs("id")%>" class="none"/>
                 </div></td>
                  <td  ><div align="center"><%=rs("id")%></div></td> 
                  <td height="40"  > 
                    <div align="center"> <%=rs("shipmentName")%></div></td>
                  <td ><div align="center"><%=rs("shipmentMoney")%></div></td>
                  <td >
				    <div align="center"><a href="psfsxg.asp?ID=<%=rs("ID")%>">修改</a>&nbsp; 
			        &nbsp;<a href="psfsdel.asp?ID=<%=rs("ID")%>" onClick="return confirm('确定要删除该配送方式吗？');" >删除</a></div></td>
                </tr>
 <%
rs.MoveNext
Next	
%>    
 <tr  >
                  <td height="40"  > 
                  
                      
                      <div align="center">
                        <input type="checkbox" name="allbox" onClick="CheckAll()" class="none"/>
                      </div></td>
                  <td colspan="4" style="padding:5px; text-align:left" ><input class="bnt" type="submit" name="button" id="button" value="批量删除"   onClick="return confirm('确定要批量删除选中的配送方式吗，将不可恢复?');"/></td>
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
	<table width="99%"  cellpadding="0" cellspacing="0" class="table1" style="margin-top:10px">
  		
          <tr   >
            <td  height="30"  style="padding-left:10px;border-bottom:none; text-align: center;">  <%Response.Write "目前暂无任何信息"
Response.end%> </td>       
  </table>
<%Set Obj_Page = Nothing

RS.Close
Set RS = Nothing
End If
%>
</form>
</body>
</html>
