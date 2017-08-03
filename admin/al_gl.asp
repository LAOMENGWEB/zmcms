<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="class_page.asp"-->
<!--#include file="Inc/Function.asp"-->
<!--#include file="htmlconfig.asp"-->
<!--#include file="admin_html_function.asp"-->
<%
if Instr(session("AdminPurview"),"|43,")=0 then 
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

If request.Form("button") = "更新时间" Then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='al_gl.asp';</script>" 
response.end()

end if
dim sqlg 
if sqlno=1 then
sqlg="update al set adddate=now() where id in ("&Request("id")&")"
else
sqlg="update al set adddate=getdate() where id in ("&Request("id")&")"
end if
conn.Execute ( sqlg )
if yy=2 then
call htmlalfl
call htmlal
call htmll("","","default"&lang&".html","default.asp","m=",mde,"language=",lang,"","","","","","","","")
call htmll("/"&photo&"/","",""&alclass&".html",""&photo&"/index.asp","m=",mph,"language=",lang,"","","","","","","","")
Response.Write "<script>window.location.href='al_gl.asp';</script>" 
end if
call addlog("更新"&tupian&"的发表时间为最新时间")
end if
If request.Form("button") = "设为推荐" Then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='al_gl.asp';</script>" 
response.end()
end if
dim sql4 
sql4="update al set tj=1 where id in ("&Request("id")&")"
conn.Execute ( sql4 )
if yy=2 then
call htmlalfl
call htmlal
call htmll("","","default"&lang&".html","default.asp","m=",mde,"language=",lang,"","","","","","","","")
call htmll("/"&photo&"/","",""&alclass&".html",""&photo&"/index.asp","m=",mph,"language=",lang,"","","","","","","","")
Response.Write "<script>window.location.href='al_gl.asp';</script>" 
end if
call addlog("设为推荐"&tupian&"")
end if
If request.Form("button") = "取消推荐" Then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='al_gl.asp';</script>" 
response.end()
end if
dim sql2 
sql2="update al set tj=0 where id in ("&Request("id")&")"
conn.Execute ( sql2 )
if yy=2 then
call htmlalfl
call htmlal
call htmll("","","default"&lang&".html","default.asp","m=",mde,"language=",lang,"","","","","","","","")
call htmll("/"&photo&"/","",""&alclass&".html",""&photo&"/index.asp","m=",mph,"language=",lang,"","","","","","","","")
Response.Write "<script>window.location.href='al_gl.asp';</script>" 
end if
call addlog("取消推荐"&tupian&"")
end if
If request.Form("button") = "批量删除" Then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='al_gl.asp';</script>" 
response.end()
end if
for y=1 to request("ID").count
if request("ID").count=1 then
ArticleID=request("ID")
else
ArticleID=replace(request("id")(y),"'","")
end if	
set rs=server.CreateObject("adodb.recordset")
sql = "select * from al where id ="&ArticleID&""
rs.open sql,conn,1,3	
pone=rs("zmone")	
zdfy=rs("zdfy")	
content=rs("content")

if Instr(rs("content"),"[zmcmsfy]")=0 then
call DeleteFile("../html/"&alName&""&Separated&""&pone&""&Separated&""&ArticleID&""&Separated&"1.html")
else
arrContent=split(rs("content"),"[zmcmsfy]")
pages=Ubound(arrContent)+1
for j=1 to pages
call DeleteFile("../html/"&alName&""&Separated&""&pone&""&Separated&""&ArticleID&""&Separated&""&j&".html")
next
end if

rs.close:set rs=Nothing
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from al where ID in ("&ArticleID&")",conn,0,1
IcoImage=rs("SmallImage")
SmallImage = rs("SmallImage")
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
conn.Execute ( "delete from al where id in ("&Request("id")&")" )

if yy=2 then
call htmlalfl
call htmlal
call htmll("","","default"&lang&".html","default.asp","m=",mde,"language=",lang,"","","","","","","","")
call htmll("/"&photo&"/","",""&alclass&".html",""&photo&"/index.asp","m=",mph,"language=",lang,"","","","","","","","")
end if
call addlog("删除"&tupian&"")
response.Write "<script language=javascript>alert(""删除成功！"");window.location.href=""al_gl.asp""</script>"
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
   if(confirm("确定要删除选中的<%=tupian%>吗？一旦删除将不能恢复！"))
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
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center"><%=tupian%>列表管理 <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>


<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="anlie.asp" target="main">返回菜单</a></div>
 <div class="dmeun"><a href="al_add.asp" target="main">添加<%=tupian%></a></div></td>
 

 </tr>
</table>
<table width="99%" align="center" cellpadding="0" cellspacing="0" class="dtable1">
  		
          <tr   >
            <td width="10%" >按照分类查找：</td>
			<td width="60%">
              <select name="ClassID" id="ClassID" onChange="javascript:window.open(this.options[this.selectedIndex].value,'main')">
    <%
   Dim Sqlp,Rsp,TempStr
   Sqlp ="select * from alclass where ParentClassID=0 and lang='"&lang&"' order by px asc"   
   Set Rsp=server.CreateObject("adodb.recordset")   
   rsp.open sqlp,conn,1,1 
   Response.Write("<option value="""">请选择分类</option>") 
   If Rsp.Eof and Rsp.Bof Then
      Response.Write("<option value="""">请先添加分类</option>")
   Else
	   Response.Write("<option value=""al_gl.asp"">所有分类信息</option>") 
      Do while not Rsp.Eof   
         Response.Write("<option value=" & """?pone=" & Rsp("ID") & """" & "")
		 If int(request("pone"))=Rsp("ID") then
				Response.Write(" selected" ) 
		 End if
         Response.Write(">|-" & Rsp("ClassName") & " </option>")  & VbCrLf
		 
		    Sqlpp ="select * from alclass Where ParentClassID="&Rsp("ID")&" and lang='"&lang&"'  order by px asc"   
   			Set Rspp=server.CreateObject("adodb.recordset")   
   			rspp.open sqlpp,conn,1,1
			Do while not Rspp.Eof 
				Response.Write("<option value=" & """?ptwo=" & Rspp("ID") & " """ & "")
				If int(request("ptwo"))=Rspp("ID") then
				Response.Write(" selected" ) 
				End if
         		Response.Write(">　|-" & Rspp("ClassName") & "")
				Response.Write("</option>" )   & VbCrLf
				
				
								    Sqlppp ="select * from alclass Where ParentClassID="&Rspp("ID")&" and lang='"&lang&"'  order by px asc"   
   			Set Rsppp=server.CreateObject("adodb.recordset")   
   			rsppp.open sqlppp,conn,1,1
			Do while not Rsppp.Eof 
				Response.Write("<option value=" & """?pthree=" & Rsppp("ID") & " """ & "")
				If int(request("pthree"))=Rsppp("ID") then
				Response.Write(" selected" ) 
				End if
         		Response.Write(">　　　|-" & Rsppp("ClassName") & "")
				Response.Write("</option>" )   & VbCrLf
				
				Rsppp.Movenext   
      		Loop
			rsppp.close
			set rsppp=nothing
				
				
				
				
			Rspp.Movenext   
      		Loop
			rspp.close
			set rspp=nothing
      Rsp.Movenext   
      Loop   
   End if
	rsp.close
	set rsp=nothing
	%>
        </select></td>
		
		<form id="form1" name="form1" method="post" action="al_gl.asp?Result=ok">
		
            <td width="8%">
             快速搜索：</td>
            <td width="22%">
                <input name="k1" type="text" id="k1" />
             
                           <input class="bnt" type="submit" name="button" id="button" value="搜索" />
          </td>
			</form>
          </tr>
          
        
</table>

<%
if Result="ok"   then
 DB_Sql = "select * from al where title like '%"&keytag&"%' and lang='"&lang&"' "
else
if pone="" then
DB_Sql = "select * from al where lang='"&lang&"' "
end if
if pone<>""  then
DB_Sql = "select * from al where zmone="&Pone&" and lang='"&lang&"' "
end if
if ptwo<>""   then
DB_Sql = "select * from al where  zmtwo="&Ptwo&" and lang='"&lang&"' "
end if
if pthree<>""   then
DB_Sql = "select * from al where  zmthree="&Pthree&" and lang='"&lang&"' "
end if
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


if Result="ok"   then
Obj_Page.Setpage_Pathurl = "al_gl.asp?Result=ok" '执行分页的页面
else

if pone="" then
Obj_Page.Setpage_Pathurl = "al_gl.asp" '执行分页的页面
end if
if pone<>"" then
Obj_Page.Setpage_Pathurl = "al_gl.asp?pone="&pone&"" '执行分页的页面
end if

if ptwo<>"" then
Obj_Page.Setpage_Pathurl = "al_gl.asp?ptwo="&ptwo&"" '执行分页的页面
end if

if pthree<>"" then
Obj_Page.Setpage_Pathurl = "al_gl.asp?pthree="&pthree&"" '执行分页的页面
end if


end if








Obj_Page.Setpage_Navtype = 1 '分页导航样式(可选1文字,2数字)

Call Obj_Page.Page_Init() '初始化分页

%>
<form name="zmqy" method="Post">
  <table width="99%"   align="center" cellpadding="0" cellspacing="0" class="dtable2">
               
    <tr class="wz">
      <td width="4%" ><div align="center">全选</div></td>
      <td width="5%" ><div align="center">编号</div></td>
      <td width="11%" ><div align="center">缩图</div></td>
      <td width="12%" ><div align="center">属性</div></td>
      <td width="12%" ><div align="center">所属分类</div></td>
      <td width="14%" ><div align="center">查看权限</div></td>
      <td width="17%" ><div align="center"><%=tupian%>名称</div></td>
      <td width="14%" ><div align="center">更新日期</div></td>
      <td width="23%" ><div align="center">基本操作</div></td>
    </tr>
		                 <%

'/循环开始
For II = 1 to Obj_PageListnum
If rs.Eof Then
	Exit For
End If
%>
   
                <tr > 
                  <td  > 
                    <div align="center">
                      <input name="ID" type="checkbox" id="ID" value="<%=rs("id")%>" style="border:none"/>
                  </div></td>
                  <td  ><div align="center"><%=rs("id")%></div></td>
                  <td  ><div align="center"><%if Rs("SmallImage")<>"" then%><img src="../<%=Rs("SmallImage")%>" width="50" height="50"><%else%><img src="../images/demo.jpg" width="50" height="50"><%end if%></div></td>
                  <td  >  <div align="center">
                    <%
if rs("tj")=1 then
response.Write("<font color=#FF0000>{推荐}</font>")
else
response.Write("<font color=#FF0000>{无}</font>")
end if

%>
                  </div></td>
                  <td  ><div align="center">
                    <%dim rsc
		  set rsc=server.CreateObject("adodb.recordset")
		  rsc.open "select * from alclass where lang='"&lang&"'",conn,1,1
		  while not rsc.eof
		    if rs("zmtwo")=0 and rs("zmthree")=0 and rs("zmone")=rsc("id") then
			response.Write("" & rsc("classname") & "")
			end if
			if rs("zmone")<>0  and rs("zmthree")=0 and  rs("zmtwo")=rsc("id") then
			response.Write("" & rsc("classname") & "")
			end if
			if rs("zmone")<>0 and rs("zmtwo")<>0 and  rs("zmthree")=rsc("id") then
			response.Write("" & rsc("classname") & "")
			end if
			rsc.movenext
		wend
		rsc.close
		set rsc=nothing
		  %>
                  </div></td>
                  <td  ><div align="center">
                    <%if rs("Exclusive")=">=" then
        Response.Write ""&ViewGroupName(rs("GroupID"))&" <font color='green'>隶</font>" & vbCrLf
      else
        Response.Write ""&ViewGroupName(rs("GroupID"))&" <font color='red'>专</font>" & vbCrLf
	  end if%>
                  </div></td>
                  <td  >
                    <div align="center">
                    
                    &nbsp;<font color="<%=rs("btcolor")%>"><%=rs("title")%></font></div></td>
                  <td  ><div align="center"><%=rs("AddDate") %></div></td>
                  <td  style="border-right:none"><div align="center"><a href="al_edit.asp?id=<%=rs("id")%>">修改</a> 
                        <a href="aldel.asp?id=<%=rs("id")%>" onClick="return ConfirmDel();">删除</a>
                     </div></td>
                </tr>
               <%
rs.MoveNext
Next
	
%>  <tr>
           <td  >
             <div align="center">
               <input type="checkbox" name="allbox" onClick="CheckAll()" style="border:none"/>
             </div></td>
            <td  colspan="9" style="padding:5px; text-align:left">
              
			  <input class="bnt" type="submit" name="button" id="button" value="设为推荐"  onClick="return confirm('确定要设为推荐吗？');"/>
			  <input class="bnt" type="submit" name="button" id="button" value="取消推荐" onClick="return confirm('确定要取消推荐吗？');" />
			  <input class="bnt" type="submit" name="button" id="button" value="更新时间" onClick="return confirm('确定要更新选中的<%=tupian%>为最新时间吗？');" />
              <input  type="submit" name="button" id="button" value="批量删除"  class="bnt" onClick="return confirm('确定要批量删除吗，将不可恢复?');"/>
&nbsp;             </td>
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
<table width="99%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td style="padding:10px 0 ;border-right:none;border-bottom:none">	<div align="center">
      <%
Response.Write "目前暂无任何信息"
Response.end
	%>
    </div></td>
  </tr>
</table>
<%
Set Obj_Page = Nothing
RS.Close
Set RS = Nothing
End If
%>         </form> 
</body>
</html>
<%
Function ViewGroupName(GruopID)
  dim rs,sql
  set rs = server.createobject("adodb.recordset")
  sql="select GroupID,GroupName from hyGroup where GroupID='"&GruopID&"' and lang='"&lang&"'"
  rs.open sql,conn,1,1
  if rs.bof and rs.eof then
    ViewGroupNameSi="未设组别"
  else
    ViewGroupName=rs("GroupName")
  end if
  rs.close
  set rs=nothing
end Function
%>
