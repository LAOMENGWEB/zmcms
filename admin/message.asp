<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="Inc/Function.asp"-->
<!--#include file="class_page.asp"-->


<%
if Instr(session("AdminPurview"),"|10411,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</script>"
  response.end
end if
'========判断是否具有管理权限
%>


<%
If request("action") = "cl" Then
dim id
id=request.QueryString("id")
Set rs = server.CreateObject("adodb.recordset")
sql = "select * from jszx where id="&id&" "
rs.Open sql, conn, 1, 3
rs("BizOK")=1
rs.update
    rs.Close
    Set rs = Nothing
call addlog("处理客户咨询信息")
end if
%>

<%
If request("action") = "qxcl" Then

id=request.QueryString("id")
Set rs = server.CreateObject("adodb.recordset")
sql = "select * from jszx where id="&id&" "
rs.Open sql, conn, 1, 3
rs("BizOK")=0
rs.update
    rs.Close
    Set rs = Nothing
call addlog("取消处理客户咨询信息")
end if
%>

<% 
sh=request.QueryString("sh")
pone=request.QueryString("lyid")
If request.Form("button") = "批量处理" Then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='message.asp';</script>" 
response.end()
end if
dim sql3 
sql3="update jszx set BizOK=1 where id in ("&Request("id")&")"
conn.Execute ( sql3 )
call addlog("处理即时咨询留言")
end if

If request.Form("button") = "取消处理" Then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='message.asp';</script>" 
response.end()
end if
dim sql4 
sql4="update jszx set BizOK=0 where id in ("&Request("id")&")"
conn.Execute (sql4)
call addlog("取消处理信息")
end if

If request.Form("button") = "批量删除" Then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='message.asp';</script>" 
response.end()
end if

for y=1 to request("ID").count
if request("ID").count=1 then
ArticleID=request("ID")
else
ArticleID=replace(request("id")(y),"'","")
end if	
set rs=server.CreateObject("adodb.recordset")
sql = "select * from jszx where id ="&ArticleID&""
rs.open sql,conn,1,3
next
sql5="delete from jszx where id in ("&Request("id")&")"
conn.Execute ( sql5 )


call addlog("删除客户及时咨询信息")
Response.Write "<script language=javascript>alert('恭喜!操作成功!');window.location.href='message.asp';</script>" 
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



<form action="message.asp" method="post" name="form1" id="form1"> 
<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">客户及时咨询管理  </div></td>
</tr>
</table>




<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td >
 <div class="dmeun"><a href="message.asp" target="main">所有信息</a> </div>
  <div class="dmeun"><a href="message.asp?sh=0" target="main">未处理信息</a> </div>
   <div class="dmeun"><a href="message.asp?sh=1" target="main">已处理信息</a> </div>
 </td>
 

 </tr>
</table>
<%




if sh<>"" or pone<>"" then
if sh=0 then
DB_Sql = "select * from jszx where BizOK=0 and lang='"&lang&"' "
end if
if sh=1 then
DB_Sql = "select * from jszx where BizOK=1 and lang='"&lang&"' "
end if

else
DB_Sql = "select * from jszx where lang='"&lang&"'  "
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
Obj_Page.Setpage_Pathurl = "message.asp?sh=0" '执行分页的页面
end if
if sh=1 then
Obj_Page.Setpage_Pathurl = "message.asp?sh=1" '执行分页的页面
end if


else

Obj_Page.Setpage_Pathurl = "message.asp" '执行分页的页面
end if





Obj_Page.Setpage_Navtype = 1 '分页导航样式(可选1文字,2数字)

Call Obj_Page.Page_Init() '初始化分页

%>
            
              <table width="99%" align="center" cellpadding="0" cellspacing="0"  class="dtable2">
               
                <tr  class="wz">
                  <td >全选</td>
                  <td  >ID</td>
                  <td   >客户咨询内容</td>
                  <td  >电话</td>
                  <td   >地址</td>
                  <td >信箱</td>
                  <td >咨询发布时间</td>
                  <td >处理状态</td>
                  <td >操作</td>
                </tr>
             <%

'/循环开始
For II = 1 to Obj_PageListnum
If rs.Eof Then
	Exit For
End If
%>
                <tr >
                  <td ><input style=" border:0px #ddd solid" name="ID" type="checkbox" id="ID" value="<%=rs("id")%>" /></td>
                  <td ><%=rs("id")%></td>
              
                  <td ><%=rs("BizContent")%></td>
                  <td ><%=rs("BizPhone")%></td>
                  <td ><%=rs("BizAddr")%></td>
                  <td ><%=rs("BizEMail")%>&nbsp;<a href="sendemail.asp?l=somail&selectemail=<%=rs("BizEMail")%>"><font color="#FF0000">回信</font></a></td>
                  <td ><%=rs("BizDate")%></td>
                  <td ><%If rs("BizOK") = 1 Then%>
   <font color='blue'>已处理</font>
   
   <%
Else%>
    <font color='red'>未处理</font>
	<%
End If%></td>
                  <td >				  
                    
                    &nbsp;<a href="message.asp?action=cl&ID=<%=rs("ID")%>">处理</a>&nbsp; 
					&nbsp;<a href="message.asp?action=qxcl&ID=<%=rs("ID")%>">取消处理</a>&nbsp; 
                    &nbsp;<a href="jssc.asp?ID=<%=rs("ID")%>" onClick="return ConfirmDel('确定要删除网站吗？');">删除</a></td>
                </tr>
<%
rs.MoveNext
Next
	
%>    
<tr>
                  <td > 
                    <div align="center">
                      <input style=" border:0px #ddd solid" type="checkbox" name="allbox" onClick="CheckAll()" />
      </div></td>
                  <td  colspan="8" style="padding:5px; text-align:left">
                    <input class="bnt" type="submit" name="button" id="button" value="批量处理" />
					<input class="bnt" type="submit" name="button" id="button" value="取消处理" />
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
	    
 <table width="99%" border="0" align="center" cellpadding="0" cellspacing="0"  class="table1" style="margin-top:10px">
  <tr>
    <td >
	 
	    <%
    
Response.Write "目前暂无任何信息"
Response.end
	%>
      </td>
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