<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="Inc/Function.asp"-->
<!--#include file="class_page.asp"-->
<%
if Instr(session("AdminPurview"),"|34,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</script>"
  response.end
end if
'========判断是否具有管理权限
%>
<% 
ddh=request.QueryString("ddh") 
sh=request.QueryString("sh")
pone=request.QueryString("orderid")
If request.Form("button") = "设为支付" Then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='order.asp';</script>" 
response.end()
end if
dim sql4 
sql4="update orderdd set isdate=1 where id in ("&Request("id")&")"
conn.Execute (sql4)
call addlog("设为支付")
end if

If request.Form("button") = "设为处理中" Then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='order.asp';</script>" 
response.end()
end if

sql4="update orderdd set isdate=2 where id in ("&Request("id")&")"
conn.Execute (sql4)
call addlog("设为处理中")
end if

If request.Form("button") = "设为已发货" Then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='order.asp';</script>" 
response.end()
end if

sql4="update orderdd set isdate=3 where id in ("&Request("id")&")"


conn.Execute (sql4)
call addlog("设为已发货")
end if

If request.Form("button") = "订单完成" Then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='order.asp';</script>" 
response.end()
end if

sql4="update orderdd set isdate=5 where id in ("&Request("id")&")"


conn.Execute (sql4)
call addlog("设为订单完成")
end if

If request.Form("button") = "批量删除" Then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='order.asp';</script>" 
response.end()
end if
for y=1 to request("ID").count
if request("ID").count=1 then
ArticleID=request("ID")
else
ArticleID=replace(request("id")(y),"'","")
end if	
set rs2=server.CreateObject("adodb.recordset")	
sql2="select * from orderdd where id in ("&Request("id")&")"
rs2.open sql2,conn,1,1
For i = 1 To rs2.RecordCount
conn.Execute ( "delete from ordercp where orderddh =('"&rs2("orderddh")&"')" )
rs2.movenext
next
set rs2=nothing
rs2.close
next
conn.Execute ( "delete from orderdd where id in ("&Request("id")&")" )
call addlog("删除订单")
response.Write "<script language=javascript>alert(""删除成功！"");window.location.href=""order.asp""</script>"
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
   if(confirm("确定要删除选中的订单吗？一旦删除将不能恢复！"))
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
<form id="form1" name="form1" method="get" action="order.asp">




  
<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">订单管理  <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>




<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td >
 <div class="dmeun"><a href="order.asp" target="main">所有订单</a></div>
 <div class="dmeun"><a href="order.asp?sh=1" target="main">已处理订单</a></div>
  <div class="dmeun"><a href="order.asp?sh=2" target="main">未处理订单</a></div>
   <div class="dmeun"><a href="order.asp?sh=3" target="main">游客订单</a></div>
   <div class="dmeun"><a href="order.asp?sh=4" target="main">会员订单</a></div>
 </td>
 

 </tr>
</table>

  
  
  
  
  
  
  
  
  <table width="99%" align="center" cellpadding="0" cellspacing="0" class="dtable1">
  <tr>
  <td width="11%" >
      
            通过订单号查询：</td>
  <td width="89%"><input type="text" name="ddh" />
         
               
                    <input type="submit" name="Submit" value="搜索" class="bnt" />
           
      </td>
  
  </tr>
  </table>
  
  
</form>
<form name="zmqy" method="Post">
  <%
  if pone="" or ddh="" then
DB_Sql = "select * from Orderdd where lang='"&lang&"' "
end if

  if pone<>"" then
DB_Sql = "select * from Orderdd where MemID="&pone&" and lang='"&lang&"' "
end if

if sh=1 then
DB_Sql = "select * from Orderdd where isdate<>0 and lang='"&lang&"' "
end if

if sh=2 then
DB_Sql = "select * from Orderdd where isdate=0 and lang='"&lang&"' "
end if

if sh=3 then
DB_Sql = "select * from Orderdd where memid=0 and lang='"&lang&"' "

end if
if sh=4 then
DB_Sql = "select * from Orderdd where memid=1 and lang='"&lang&"' "

end if

if ddh<>"" then
DB_Sql = "select * from Orderdd where orderddh='"&ddh&"' and lang='"&lang&"'"
end if
if orderid<>"" then
DB_Sql = "select * from Orderdd where orderid='"&pone&"' and lang='"&lang&"'"
end if

DB_Sql = DB_Sql & "order by  ID desc"
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

if pone="" or ddh="" then
Obj_Page.Setpage_Pathurl = "order.asp" '执行分页的页面
end if

if sh=1 then
Obj_Page.Setpage_Pathurl = "order.asp?sh=1" '执行分页的页面
end if

if sh=2 then
Obj_Page.Setpage_Pathurl = "order.asp?sh=2" '执行分页的页面
end if

if sh=3 then
Obj_Page.Setpage_Pathurl = "order.asp?sh=3" '执行分页的页面
end if

if sh=4 then
Obj_Page.Setpage_Pathurl = "order.asp?sh=4" '执行分页的页面
end if







Obj_Page.Setpage_Navtype = 1 '分页导航样式(可选1文字,2数字)

Call Obj_Page.Page_Init() '初始化分页
%>
  <table width="99%" align="center" cellpadding="0" cellspacing="0" class="dtable2">
  		
          <tr  class="wz">
            <td  ><div align="center">全选</div></td>
            <td  ><div align="center">ID</div></td>
            <td  ><div align="center">订单号码</div></td>
            <td  ><div align="center">订购人</div></td>
            <td  ><div align="center">产品订购</div></td>
            <td ><div align="center">订购单位</div></td>
            <td ><div align="center">订购时间</div></td>
            <td ><div align="center">订单详情</div></td>
            <td ><div align="center">查看详情</div></td>
            <td ><div align="center" >基本操作</div></td>
          </tr>
<%
'/循环开始
For II = 1 to Obj_PageListnum
If rs.Eof Then
	Exit For
End If
%>
          <tr > 
            <td width="4%" > 
              <div align="center">
  <input name="ID" type="checkbox" id="ID" value="<%=rs("id")%>" class="none"/>
 </div></td>
		    <td width="2%"><div align="center"><%=rs("id")%></div></td>
		    <td width="13%"><div align="center"><%=rs("orderddh")%></div></td>
		    <td width="13%"><div align="center"><%=Guest(rs("MemID"),rs("Linkman"))%></div></td>
		    <td width="7%"><div align="center"><%=rs("OrderName")%></div></td>
		    <td width="12%"><div align="center"><%=rs("Company")%></div></td>
		    <td width="14%"><div align="center"><%=rs("AddTime")%></div></td>
		    <td width="14%"> <div align="center">
		<%
	 If rs("isdate") = 0 Then
      response.write("等待付款")
	  end if
	   If rs("isdate") = 1 Then
      response.write("已支付")
	  end if
	   If rs("isdate") = 2 Then
      response.write("处理中")
	  end if
	   If rs("isdate") = 3 Then
      response.write("已发货")
	  end if
	   If rs("isdate") = 5 Then
      response.write("订单完成")
	  end if
	  %>
	  </div></td>
		    <td width="11%"><div align="center"><a href="ddxx.asp?Result=Modify&ID=<%=rs("id")%>">查看订单详情</a></div></td>
		    <td width="10%" style="border-right:none"><div align="center" ><a href="ddsc.asp?id=<%=rs("orderddh")%>" onClick="return ConfirmDel();">删除</a></div></td>
          </tr>
		      <%
rs.MoveNext
Next	
%>          <tr    >
<td><input type="checkbox" name="allbox" onClick="CheckAll()" class="none"/></td>
           <td  colspan="10" style="padding:5px; text-align:left">
             <input class="bnt" type="submit" name="button" id="button" value="批量删除"   onClick="return confirm('确定要批量删除选中的订单信息吗，将不可恢复?');"/>
			1. <input class="bnt" type="submit" name="button" id="button" value="设为支付"   onClick="return confirm('确定要设为已支付吗?');"/>
			2. <input class="bnt" type="submit" name="button" id="button" value="设为处理中"   onClick="return confirm('确定要设为处理中吗?');"/>
			3. <input class="bnt" type="submit" name="button" id="button" value="设为已发货"   onClick="return confirm('确定要设为已发货吗?');"/>
			
			4.等待客户确定收货后，订单自动完成，也可以手动设置订单完成
			
			<input class="bnt" type="submit" name="button" id="button" value="订单完成"   onClick="return confirm('确定要设为订单完成吗?');"/>
&nbsp;          </td>
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
 <table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" class="table1">
  <tr>
    <td style="border-bottom:none;border-right:none">
	  <div align="center">
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

<%
function Guest(ID,Linkman)
  Dim rs,sql
  Set rs=server.CreateObject("adodb.recordset")
  sql="Select * From Members where ID="&ID
  rs.open sql,conn,1,1
  if rs.bof and rs.eof then
    Guest="{未注册用户}"&Linkman&""
  else
    Guest="<font color='red'>{"&rs("GroupName")&"}</font><a href='Memberxg.asp?Result=Modify&ID="&ID&"'>"&Linkman&"</a>"
  end if
  rs.close
  set rs=nothing
end function
%>
</body>
</html>
