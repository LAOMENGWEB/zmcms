<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="class_page.asp"-->
<!--#include file="Inc/Function.asp"-->
<!--#include file="htmlconfig.asp"-->
<!--#include file="admin_html_function.asp"-->
<%
if Instr(session("AdminPurview"),"|72,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</script>"
  response.end
end if
'========判断是否具有管理权限
%>
<% 
If request.Form("button") = "批量删除" Then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='zp.asp';</script>" 
response.end()
end if





'批量删除开始
for y=1 to request("ID").count
if request("ID").count=1 then
ArticleID=request("ID")
else
ArticleID=replace(request("id")(y),"'","")
end if	
set rs=server.CreateObject("adodb.recordset")
sql = "select * from job where id ="&ArticleID&""
rs.open sql,conn,1,3	
pone=rs("zmone")
zdfy=rs("zdfy")	
content=rs("HrDetail")

if Instr(rs("HrDetail"),"[zmcmsfy]")=0 then
call DeleteFile("../html/"&zpName&""&Separated&""&ArticleID&""&Separated&"1.html")
else
arrContent=split(rs("HrDetail"),"[zmcmsfy]")
pages=Ubound(arrContent)+1
for j=1 to pages
call DeleteFile("../html/"&zpName&""&Separated&""&ArticleID&""&Separated&""&j&".html")
next
end if

rs.close:set rs=Nothing

'批量删除图片
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from job where ID in ("&ArticleID&")",conn,0,1
IcoImage=rs("SmallImage")
SmallImage = rs("SmallImage")
dim picUrl
dim picUrlArray
dim x,y
picUrl = getPicUrl(rs("HrDetail"))
if picUrl <> "" then
picUrl = left(picUrl,len(picUrl)-1)
picUrlArray = split(picUrl,",")
for x = 0 to ubound(picUrlArray)
if instr(picUrlArray(x),"") > 0 then
call DeleteFile(picUrlArray(x))
end if
next
end if


dim picUrl1
dim picUrlArray1
dim x1,y1
picUrl1 = rs("ImagePath")
if picUrl1 <> "" then
picUrlArray1 = split(picUrl1,"|")
for x1 = 0 to ubound(picUrlArray1)
call DeleteFile(picUrlArray1(x1))
next
end if
call deletefile1(SmallImage)
call deletefile(IcoImage)
next

rs.close:set rs=Nothing
'结束












conn.Execute ( "delete from job  where id in ("&Request("id")&")" )

rs.close
set rs=nothing
if yy=2 then
call htmlzpfl
call htmlzp
call htmll("","","default"&lang&".html","default.asp","m=",mde,"language=",lang,"","","","","","","","")
call htmll("/"&job&"/","",""&zpclass&".html",""&job&"/index.asp","m=",mzp,"language=",lang,"","","","","","","","")
end if

call addlog("删除"&zhaopin&"信息")
Response.Write "<script>alert('恭喜!操作成功!');window.location.href='zp.asp';</script>" 

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

<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center"><%=zhaopin%>管理 <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>


<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="zhaopin.asp" target="main">返回菜单</a></div><div class="dmeun"><a href="zpadd.asp" target="main">发布<%=zhaopin%></a></div></td>
 

 </tr>
</table>

<%
DB_Sql = "select * from job where lang='"&lang&"' "
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

Obj_Page.Setpage_Pathurl = "zp.asp" '执行分页的页面
Obj_Page.Setpage_Navtype = 1 '分页导航样式(可选1文字,2数字)

Call Obj_Page.Page_Init() '初始化分页

%>


<form action="zp.asp" method="post" name="form1" id="form1" class="checkform"> 

              <table width="99%"  align="center" cellpadding="0" cellspacing="0"  class="dtable2">
               
                <tr class="wz"> 
                  <td width="4%"  > <div align="center">全选</div></td>
                  <td width="5%" ><div align="center">编号</div></td>
                  <td width="12%" > <div align="center"><%=zhaopin%>职位</div></td>
                  <td width="5%" > <div align="center"><%=zhaopin%>人数</div></td>
                  <td width="14%" ><div align="center">工作地点</div></td>
                  <td width="7%" ><div align="center">工资待遇</div></td>
                  <td width="14%" > <div align="center">发布时间</div></td>
                  <td width="14%" > <div align="center">状态</div></td>
                  <td colspan="3" ><div align="center">操　作</div></td>
                </tr>
          <%

'/循环开始
For II = 1 to Obj_PageListnum
If rs.Eof Then
	Exit For
End If
%>
                <tr > 
                  <td  > <div align="center">
                    <input name="ID" type="checkbox" id="ID" value="<%=rs("id")%>" class="none"/>
                  </div></td>
                  <td ><div align="center"><%=rs("ID")%></div></td>
                  <td ><div align="center"><font color="<%=rs("btcolor")%>"><%=rs("HrName")%></font></div>                  </td>
                  <td > <div align="center"><%=rs("HrRequireNum")%>人</div></td>
                  <td ><div align="center"><%=rs("HrAddress")%></div></td>
                  <td ><div align="center"><%=rs("HrSalary")%></div></td>
                  <td > <div align="center"><%=rs("ksdate")%></div></td>
                  <td > <div align="center"> <% if rs("jsDate")>now() then
	    response.write "   "&zhaopin&"中"
	  else
	    response.write "   已结束"
	  end if%></div></td>
                  <td width="8%" > <div align="center"><a href="zpEdit.asp?id=<%=rs("id")%>">修改</a></div></td>
                  <td width="9%"> <div align="center"><a href="rcsc.asp?id=<%=rs("id")%>" onClick="return ConfirmDel('确定要删除此人才信息吗？');">删除</a></div></td>
                 
                </tr>
  <%
rs.MoveNext
Next
	
%>      
                <tr > 
                  <td  align="center" ><input type="checkbox" name="allbox" onClick="CheckAll()" class="none"/></td>
                  <td  colspan="10" align="left" style="padding:5px; text-align:left"><input class="bnt" type="submit" name="button" id="button" value="批量删除"  onClick="return confirm('确定要批量删除选中的<%=zhaopin%>信息吗？');"/></td>
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
<table width="99%" border="0" align="center" cellpadding="0" cellspacing="0" class="dtable1">
  <tr>
    <td >	<div align="center">
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
