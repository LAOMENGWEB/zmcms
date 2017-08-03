<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="Inc/Function.asp"-->
<!--#include file="class_page.asp"-->


<%
if Instr(session("AdminPurview"),"|81,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</script>"
  response.end
end if
'========判断是否具有管理权限
%>
<% 
sh=request.QueryString("sh")
pone=request.QueryString("lyid")
If request.Form("button") = "批量审核" Then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='ly.asp';</script>" 
response.end()
end if
dim sql3 
sql3="update ly set sh=1 where id in ("&Request("id")&")"
conn.Execute ( sql3 )
call addlog("审核留言")
end if

If request.Form("button") = "取消审核" Then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='ly.asp';</script>" 
response.end()
end if
dim sql4 
sql4="update ly set sh=0 where id in ("&Request("id")&")"
conn.Execute (sql4)
call addlog("取消审核留言")
end if

If request.Form("button") = "批量删除" Then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='ly.asp';</script>" 
response.end()
end if

for y=1 to request("ID").count
if request("ID").count=1 then
ArticleID=request("ID")
else
ArticleID=replace(request("id")(y),"'","")
end if	
set rs=server.CreateObject("adodb.recordset")
sql = "select * from ly where id ="&ArticleID&""
rs.open sql,conn,1,3
next
sql5="delete from ly where id in ("&Request("id")&")"
conn.Execute ( sql5 )


call addlog("删除留言")
Response.Write "<script language=javascript>alert('恭喜!操作成功!');window.location.href='ly.asp';</script>" 
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

<script language="javascript"> 
<!-- 
function CheckAll(){ 
 for (var i=0;i<eval(form1.elements.length);i++){ 
  var e=form1.elements[i]; 
  if (e.name!="allbox") e.checked=form1.allbox.checked; 
 } 
} 
--> 
</script> 
<script language=javascript>
function ConfirmDel()
{
   if(confirm("确定要删除此记录吗？"))
     return true;
   else
     return false;
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

  
  
  


<form action="ly.asp" method="post" name="form1" id="form1"> 
    

<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">留言管理  <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>




<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td >
 <div class="dmeun"><a href="ly.asp" target="main">所有留言</a></div>
 <div class="dmeun"><a href="ly.asp?sh=0" target="main">未审核留言</a></div>
  <div class="dmeun"><a href="ly.asp?sh=1" target="main">已审核留言</a></div>
   <div class="dmeun"><a href="ly.asp?sh=2" target="main">已回复留言</a></div>
    <div class="dmeun"><a href="ly.asp?sh=3" target="main">未回复留言</a></div>
 </td>
 

 </tr>
<%

if sh<>"" or pone<>"" then
if sh=0 then
DB_Sql = "select * from ly where sh=0 and lang='"&lang&"' "
end if
if sh=1 then
DB_Sql = "select * from ly where sh=1 and lang='"&lang&"' "
end if
if sh=2 then
DB_Sql = "select * from ly where hf is not null and lang='"&lang&"' "
end if

if sh=3 then
DB_Sql = "select * from ly where hf is null and lang='"&lang&"' "
end if
if pone<>""  then
DB_Sql = "select * from ly where MemID="&pone&" and lang='"&lang&"' "
end if
else
DB_Sql = "select * from ly where lang='"&lang&"'  "
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
if sh<>"" or pone<>"" then

if sh=0 then
Obj_Page.Setpage_Pathurl = "ly.asp?sh=0" '执行分页的页面
end if
if sh=1 then
Obj_Page.Setpage_Pathurl = "ly.asp?sh=1" '执行分页的页面
end if
if sh=2 then
Obj_Page.Setpage_Pathurl = "ly.asp?sh=2" '执行分页的页面
end if
if sh=3 then
Obj_Page.Setpage_Pathurl = "ly.asp?sh=3" '执行分页的页面
end if
if pone<>""  then
Obj_Page.Setpage_Pathurl = "ly.asp?lyid="&pone&"" '执行分页的页面
end if
else
Obj_Page.Setpage_Pathurl = "ly.asp" '执行分页的页面

end if
Obj_Page.Setpage_Navtype = 1 '分页导航样式(可选1文字,2数字)

Call Obj_Page.Page_Init() '初始化分页

%>

 
</table>
            
              <table width="99%" align="center" cellpadding="0" cellspacing="0" class="dtable2">
               
                <tr class="wz" >
                  <td>全选</td>
                  <td >用户</td>
                  <td >QQ</td>
                  <td >邮箱</td>
                  <td >电话</td>
                  <td >状态</td>
                  <td >留言时间</td>
                  <td >回复时间</td>
                  <td  >审核状态</td>
                  <td > 操作</td>
                </tr>
             <%

'/循环开始
For II = 1 to Obj_PageListnum
If rs.Eof Then
	Exit For
End If
%>
                <tr >
                  <td ><input  name="ID" type="checkbox" id="ID" value="<%=rs("id")%>"  class="none"/></td>
				  
                  <td ><%=Guest(rs("MemID"),rs("title"))%></td>
              
                  <td ><% if rs("qq")<>""  then %>
                <a href=tencent://message/?uin=<%=rs("qq")%>&Site=yes target="_blank"><img src="image/qq.gif" alt="<%=rs("qq")%>" width=16 height=16 border="0"></a>
<% end if %></td>
                  <td ><%=rs("emaill")%> <a href="sendemail.asp?l=somail&selectemail=<%=rs("emaill")%>"><font color="#FF0000">回信</font></a></td>
                  <td ><%=rs("phone")%></td>
                  <td ><%if rs("qqh")=1 then
        Response.Write "<font color='red'>悄悄话</font>" & vbCrLf
      else
        Response.Write "普通" & vbCrLf
	  end if%></td>
                  <td ><%=rs("sj")%></td>
                  <td ><%
if IsNull(rs("hf")) or trim(rs("hf")&"")="" then
response.Write("<font color=#FF0000>此留言暂未回复！</font>")
else
%>
<%=rs("ReplyTime")%>
<%
end if
  %></td>
                  <td><%if rs("sh")=1 then
			  response.Write("[已审核]")
			  else
			  response.Write("<font color=#FF0000>[未审核]</font>")
			  end if%></td>
                  <td >				  
                    
                    &nbsp;<a href="lyModify.asp?ID=<%=rs("ID")%>">回复</a>&nbsp; 
                    &nbsp;<a href="lysc.asp?ID=<%=rs("ID")%>" onClick="return ConfirmDel('确定要删除网站吗？');">删除</a></td>
                </tr>
<%
rs.MoveNext
Next
	
%>    
<tr >
                  <td > 
                  
                      <input  type="checkbox" name="allbox" onClick="CheckAll()" class="none"/>
  </td>
                  <td style="padding:5px;  text-align:left " colspan="9">
                    <input class="bnt" type="submit" name="button" id="button" value="批量审核" />
					<input class="bnt" type="submit" name="button" id="button" value="取消审核" />
					<input class="bnt" type="submit" name="button" id="button" value="批量删除" />

&nbsp;      </td></tr>
  </table>
  <%If Obj_Page.OutErrInfo <> "" Then
		Response.Write Obj_Page.OutErrInfo '报错信息显示
		Response.End
		else
	Call Obj_Page.Page_Navigation() '分页导航
	end if
	else
	%>

  	  <table width="99%" border="0" cellpadding="0" cellspacing="0" align="center"   class="dtable" style="margin-top:10px">
<tr>
 <td >目前暂无任何信息</td>
</tr>
</table>


</form>
<%
Set Obj_Page = Nothing
RS.Close
Set RS = Nothing
End If
conn.close
set conn=nothing
%>
</body>
</html>
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