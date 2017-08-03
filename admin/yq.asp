<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="Inc/Function.asp"-->
<!--#include file="class_page.asp"-->
<%
if Instr(session("AdminPurview"),"|113,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</script>"
  response.end
end if
'========判断是否具有管理权限
%>
<% 
Result=request.QueryString("Result")
keytag=request.Form("k1")
pone=request.QueryString("pone")
ptwo=request.QueryString("ptwo")
pthree=request.QueryString("pthree")
If request.Form("button") = "批量删除" Then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='yq.asp';</script>" 
response.end()
end if
dim id
ID=trim(request("ID"))
ID1=split(ID,",")
for i=0 to ubound(ID1)
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from yqlj where ID in ("&id1(i) &")",conn,0,1
call DeleteFile(rs("logoUrl"))
next
conn.Execute ( "delete from yqlj where id in ("&Request("id")&")" )
call addlog("删除链接")
response.Write "<script language=javascript>alert(""删除成功！"");window.location.href=""yq.asp""</script>"
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

<SCRIPT language=javascript>

function ConfirmDel()
{
   if(confirm("确定要删除吗？一旦删除将不能恢复！"))
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
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">链接管理 <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>


<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="lianjie.asp" target="main">返回菜单</a></div><div class="dmeun"><a href="yqzdd.asp" target="main">添加链接</a></div></td>
 

 </tr>
</table>
<table width="99%" align="center" cellpadding="0" cellspacing="0" class="dtable1">
  		
          <tr>
            <td width="10%" >按照分类查找：</td>
            <td width="90%"><select name="ClassID" id="ClassID" onChange="javascript:window.open(this.options[this.selectedIndex].value,'main')">
    <%
   Dim Sqlp,Rsp,TempStr
   Sqlp ="select * from yqljclass where ParentClassID=0 order by px asc"   
   Set Rsp=server.CreateObject("adodb.recordset")   
   rsp.open sqlp,conn,1,1 
   Response.Write("<option value="""">请选择分类</option>") 
   If Rsp.Eof and Rsp.Bof Then
      Response.Write("<option value="""">请先添加分类</option>")
   Else
   Response.Write("<option value=""yq.asp"">所有分类链接</option>") 
      Do while not Rsp.Eof   
         Response.Write("<option value=" & """?pone=" & Rsp("ID") & """" & "")
		 If int(request("pone"))=Rsp("ID") then
				Response.Write(" selected" ) 
		 End if
         Response.Write(">" & Rsp("ClassName") & " </option>")  & VbCrLf
		 
      Rsp.Movenext   
      Loop   
   End if
	rsp.close
	set rsp=nothing
	%>
        </select></td>
  
          </tr>
       
  </table>
<%

if pone="" then
DB_Sql = "select * from yqlj where lang='"&lang&"' "
end if
if pone<>""  then
DB_Sql = "select * from yqlj where zmone="&Pone&" and lang='"&lang&"'"
end if


DB_Sql = DB_Sql & "order by px asc"
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

Obj_Page.Setpage_Pathurl = "yq.asp" '执行分页的页面
Obj_Page.Setpage_Navtype = 1 '分页导航样式(可选1文字,2数字)

Call Obj_Page.Page_Init() '初始化分页

%>
<form action="yq.asp" method="post" name="form1" id="form1"> 



<table width="99%"  align="center" cellpadding="0" cellspacing="0" class="dtable2">
  <tr class="wz">
    <td width="50" align="center"  >选中</td>
    <td width="53" align="center"  >编号</td>
    <td width="151" align="center"  >排序</td>
    <td width="105" align="center"  >所属分类</td>
    <td width="168" align="center"  ><div align="center" >链接地址</div></td>
    <td width="168" align="center"  >链接名称</td>
    <td width="260" align="center"  >链接图片</td>
    <td width="124" align="center" >管理操作</td>
  </tr>
 <%

'/循环开始
For II = 1 to Obj_PageListnum
If rs.Eof Then
	Exit For
End If
%>
  <tr align="center">
    <td ><input name="ID" type="checkbox" id="ID" value="<%=rs("id")%>" class="none"/></td>
    <td ><%=rs("id")%></td>
    <td ><iframe name="zmcms" width="100" height="50" frameborder="0" scrolling="no" 
				src="xgyqpx.asp?id=<%=rs("id")%>" ></iframe></td>
    <td ><%
		  dim rsc
		  set rsc=server.CreateObject("adodb.recordset")
		  rsc.open "select * from yqljclass",conn,1,1
		  while not rsc.eof
		    if rs("zmone")=rsc("id") then
			response.Write("" & rsc("classname") & "")
			end if
			rsc.movenext
		wend
		rsc.close
		set rsc=nothing
		  %></td>
    <td ><%=rs("dz")%></td>
    <td ><%=rs("mc")%></td>
    <td >
	<%if rs("logourl")<>"" then%>
      <img src="../<%=rs("logourl")%>" width="80" height="30" alt="<%=rs("mc")%>" />
	  <%else
	  
	  response.Write("无链接图片")
	  %>
      <%end if%></td>
    <td style="border-right:none">[<a href="yqxg.asp?id=<%=rs("ID")%>">修改</a>] [<a href="yqsc.asp?id=<%=rs("ID")%>" onClick="return confirm('确定要删除网站吗？')">删除</a>]</td>
  </tr>

<%
rs.MoveNext
Next
	
%>       
<tr >
                  <td  align="center" > 
                    <div align="center">
                      <input type="checkbox" name="allbox" onClick="CheckAll()" class="none"/>
      </div></td>
      <td colspan="7" style="padding:5px; text-align:left"><input class="bnt" type="submit" name="button" id="button" value="批量删除"  onClick="return confirm('确定要批量删除选中的链接吗？');"/></td>
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
    <td  style="padding:5px 0; border-right:none; border-bottom:none">   <div align="center">
      <%
Response.Write "目前暂无任何信息"
Response.end
	%>
    </div></td>
  </tr>
</table>

	
	
	
	<%
	End If
	Set Obj_Page = Nothing

RS.Close
Set RS = Nothing

%>
<br>
</form>
</body>
</html>
