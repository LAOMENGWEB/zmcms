<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="Inc/Function.asp"-->
<!--#include file="class_page.asp"-->

<%
if Instr(session("AdminPurview"),"|1002,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</script>"
  response.end
end if
'========判断是否具有管理权限
%>
<% 
sh=request.QueryString("sh")
pone=request.QueryString("pone")
Result=request.QueryString("Result")
keytag=request.Form("k1")

If request.Form("button") = "会员开启" Then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='member.asp';</script>" 
response.end()
end if
dim sql4 
sql4="update members set working=1 where id in ("&Request("id")&")"
conn.Execute ( sql4 )
call addlog("开启会员")
end if
If request.Form("button") = "会员锁定" Then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='members.asp';</script>" 
response.end()
end if
dim sql2 
sql2="update members set working=0 where id in ("&Request("id")&")"
conn.Execute ( sql2 )
call addlog("会员锁定")
end if

If request.Form("button") = "批量审核" Then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='members.asp';</script>" 
response.end()
end if
dim sql6 
sql6="update members set sh=1 where id in ("&Request("id")&")"
conn.Execute ( sql6 )

dim sql66
sql66="update members set GroupID=1777680 where id in ("&Request("id")&")"
conn.Execute ( sql66 )

dim sql666
sql666="update members set Groupname=""注册会员"" where id in ("&Request("id")&")"
conn.Execute ( sql666 )

call addlog("审核会员")
end if

If request.Form("button") = "取消审核" Then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='members.asp';</script>" 
response.end()
end if
dim sql7 
sql7="update members set sh=0 where id in ("&Request("id")&")"
conn.Execute ( sql7 )

dim sql77 
sql77="update members set GroupID=999177 where id in ("&Request("id")&")"
conn.Execute ( sql77 )


dim sql777 
sql777="update members set Groupname=""临时游客"" where id in ("&Request("id")&")"
conn.Execute ( sql777 )




call addlog("取消审核会员")
end if




If request.Form("button") = "批量删除" Then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='members.asp';</script>" 
response.end()
end if

for y=1 to request("ID").count
if request("ID").count=1 then
ArticleID=request("ID")
else
ArticleID=replace(request("id")(y),"'","")
end if	




set rsdy=server.CreateObject("ADODB.RECORDSET")
sqldy="select * from ly where MemID="&ArticleID&" "
rsdy.open sqldy,conn,1,1
if not rsdy.eof then
response.Write "<script language=javascript>alert(""请先删除该会员发布的留言"");window.history.back();</script>"
response.End
end if
rsdy.close
set rsdy=nothing
set rsdy=server.CreateObject("ADODB.RECORDSET")
sqldy="select * from HrDemandAccept where MemID="&ArticleID&" "
rsdy.open sqldy,conn,1,1
if not rsdy.eof then
response.Write "<script language=javascript>alert(""请先删除该会员的下的应聘信息"");window.history.back();</script>"
response.End
end if
rsdy.close
set rsdy=nothing

set rsdy=server.CreateObject("ADODB.RECORDSET")
sqldy="select * from Orderdd where MemID="&ArticleID&" "
rsdy.open sqldy,conn,1,1
if not rsdy.eof then
response.Write "<script language=javascript>alert(""请先删除该会员的下的订单信息"");window.history.back();</script>"
response.End
end if
rsdy.close
set rsdy=nothing


next

conn.Execute ( "delete from members where id in ("&Request("id")&")" )
call addlog("删除会员")
response.Write "<script language=javascript>alert(""删除成功！"");window.location.href=""member.asp""</script>"
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
   if(confirm("确定要删除选中的会员吗？一旦删除将不能恢复！"))
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

<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">会员管理 <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>


<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="huiyuan.asp" target="main">返回菜单</a></div>
 <div class="dmeun"><a href="member.asp" target="main">所有会员</a></div>
 <div class="dmeun"><a href="member.asp?sh=0" target="main">未审核会员</a></div>
 <div class="dmeun"><a href="member.asp?sh=1" target="main">已审核会员</a></div>
 
 </td>
 

 </tr>
</table>
<table width="99%" align="center" cellpadding="0" cellspacing="0" class="dtable1">
  		
          <tr   >
            <td width="12%" >根据会员组别查询：</td>
			<td width="54%">
              <select name="ClassID" id="ClassID" onChange="javascript:window.open(this.options[this.selectedIndex].value,'main')">
    <%
   Dim Sqlp,Rsp,TempStr
   Sqlp ="select * from hyGroup where lang='"&lang&"'"   
   Set Rsp=server.CreateObject("adodb.recordset")   
   rsp.open sqlp,conn,1,1 
   Response.Write("<option value="""">请选择分组</option>") 
   If Rsp.Eof and Rsp.Bof Then
      Response.Write("<option value="""">请先添加分组</option>")
   Else
	   Response.Write("<option value=""member.asp"">所有会员组</option>") 
      Do while not Rsp.Eof   
         Response.Write("<option value=" & """?pone=" & Rsp("GroupID") & """" & "")
		 If request("pone")=Rsp("GroupID") then
				Response.Write(" selected" ) 
		 End if
         Response.Write(">" & Rsp("Groupname") & " </option>")  & VbCrLf
		 
      Rsp.Movenext   
      Loop   
   End if
	rsp.close
	set rsp=nothing
	%>
        </select></td>
        <form id="form1" name="form1" method="post" action="member.asp?Result=ok">
            <td width="12%" >
			
			
			
			
			
             通过会员用户名查询：</td>
			 
			 <td width="22%">
                <input name="k1" type="text" id="k1" />
             
                            <label><input class="bnt" type="submit" name="button" id="button" value="搜索" /></label>
            
		  </td>
            </form>
          </tr>
  </table>
<form name="zmqy" method="Post">
  <%

  if Result="ok"   then

 DB_Sql = "select * from members where MemName like '%"&keytag&"%' and lang='"&lang&"' "
else

if sh=0 then
DB_Sql = "select * from members where sh=0 and lang='"&lang&"' "
end if
if sh=1 then
DB_Sql = "select * from members where sh=1 and lang='"&lang&"' "
end if
if sh=""  then
DB_Sql = "select * from members where lang='"&lang&"' "
end if
if pone<>""  then
DB_Sql = "select * from members where groupid='"&pone&"' and lang='"&lang&"'"
end if

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

if Result="ok"   then
Obj_Page.Setpage_Pathurl = "member.asp?Result=ok" '执行分页的页面
else


if sh="" then
Obj_Page.Setpage_Pathurl = "member.asp" '执行分页的页面
end if
if sh=0 then
Obj_Page.Setpage_Pathurl = "member.asp?sh=0" '执行分页的页面
end if
if sh=1 then
Obj_Page.Setpage_Pathurl = "member.asp?sh=1" '执行分页的页面
end if
if pone<>"" then
Obj_Page.Setpage_Pathurl = "member.asp?pone="&pone&"" '执行分页的页面
end if

end if
Obj_Page.Setpage_Navtype = 1 '分页导航样式(可选1文字,2数字)

Call Obj_Page.Page_Init() '初始化分页
%>
  <table width="99%" align="center" cellpadding="0" cellspacing="0" class="dtable2">
  		
          <tr    class="wz">
            <td ><div align="center">全选</div></td>
            <td  ><div align="center">ID</div></td>
            <td  ><div align="center">是否启用</div></td>
            <td  ><div align="center">用户名</div></td>
            <td  ><div align="center">所属会员组</div></td>
            <td  ><div align="center">注册时间</div></td>
            <td  ><div align="center">最近登录时间</div></td>
            <td  ><div align="center">登录次数</div></td>
            <td  ><div align="center">审核状态</div></td>
            <td  ><div align="center">所属会员信息</div></td>
            <td >是否绑定QQ</td>
            <td ><div align="center">基本操作</div></td>
          </tr>
<%
'/循环开始
For II = 1 to Obj_PageListnum
If rs.Eof Then
	Exit For
End If
%>
          <tr  > 
            <td width="3%" > 
              <div align="center">
  <input name="ID" type="checkbox" id="ID" value="<%=rs("id")%>" class="none"/>
 </div></td>
		    <td width="4%"><div align="center"><%=rs("id")%></div></td>
		    <td width="6%"><div align="center">
		      <%if rs("working") then
        Response.Write "<font color='blue'>正常</font>" & vbCrLf
      else
        Response.Write "<font color='red'>已锁定</font>" & vbCrLf
	  end If%>
	        </div></td>
		    <td width="5%"><div align="center"><%=rs("MemName")%></div></td>
		    <td width="7%"><div align="center"><%=rs("GroupName")%></div></td>
		    <td width="7%"><div align="center"><%=rs("AddTime")%></div></td>
		    <td width="14%"> <div align="center">
		      <%If rs("LastLoginTime") <> "" Then%>
              <%=rs("LastLoginTime")%>
              <%
	  Else
	  %>
		                <font color=""#CC0000"">暂未登录</font>
              <%
	  End If%>
	        </div></td>
		    <td width="7%"><div align="center"><%=rs("LoginTimes")%></div></td>
		    <td width="7%"><div align="center">
		      <%if rs("sh")=0 then  response.Write("<font color=red>[未审核]</font>") else response.Write("[已审核]") end if%>
	        </div></td>
		    <td width="12%"><div align="center"><a href="ly.asp?lyid=<%=rs("id")%>">留言</a>&nbsp;<a href="yp.asp?ypid=<%=rs("id")%>">应聘</a>&nbsp;<a href="order.asp?orderid=<%=rs("id")%>">订单</a></div></td>
		    <td width="11%" ><%if rs("qqopenapi")<>"" then response.write("已绑定QQ") else response.write("未绑定QQ")  end if%></td>
		    <td width="17%" ><div align="center"><a href="Memberxg.asp?Result=Modify&ID=<%=rs("id")%>">修改</a>   <a href="Membersc.asp?ID=<%=rs("id")%>">删除</a></div></td>
          </tr>
		      <%
rs.MoveNext
Next	
%>          <tr    >
           <td align="center" ><input type="checkbox" name="allbox" onClick="CheckAll()" class="none"/> </td>
           <td colspan="12" style="padding:5px; text-align:left"><input class="bnt" type="submit" name="button" id="button" value="会员开启"  onClick="return confirm('确定要启用该会员账号吗？');"/>
			  <input class="bnt" type="submit" name="button" id="button" value="会员锁定" onClick="return confirm('确定要锁定该会员账号吗？');" />
			  <input class="bnt" type="submit" name="button" id="button" value="批量删除"   onClick="return confirm('确定要批量删除选中的会员吗，将不可恢复?');"/>
			  <input class="bnt" type="submit" name="button" id="button" value="批量审核"   onClick="return confirm('确定要批量审核选中的会员吗?');"/>
			  <input class="bnt" type="submit" name="button" id="button" value="取消审核"   onClick="return confirm('确定要取消审核选中的会员吗?');"/></td>
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
    <td style="border-bottom:0px;border-right:0px">  <div align="center">
      <%
Response.Write "目前暂无任何信息"
Response.end%>
    </div></td>
  </tr>
</table>

	%>
<%Set Obj_Page = Nothing

RS.Close
Set RS = Nothing
End If
%>
 
</form>
</body>
</html>
